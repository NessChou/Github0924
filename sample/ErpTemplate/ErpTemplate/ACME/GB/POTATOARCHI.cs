using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;
using System.Net;
namespace ACME
{
    public partial class POTATOARCHI : Form
    {
        public string c;
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        private SerialPort comport = new SerialPort();
        StringBuilder sb = new StringBuilder();
        System.Data.DataTable dtGetAcmeStageG = null;
        string MESS = "";
        string CrLf = "\r\n";
        StreamWriter sw;
        string FirmNo = "22468373";
        string PROD = "";
      //  string host = "61.57.227.80";
        string host = "ftp.bpscm.com.tw";
     
        //ftp.bpscm.com.tw
        string username = "22468373p";
        string password = "b152224$P";
        string UpLoadDataPath = "/Upload/";

        public POTATOARCHI()
        {
            InitializeComponent();
        }

        private void POTATOARCHI_Load(object sender, EventArgs e)
        {

       
            //btnPrintTest.Enabled = false;
            //if (comport.IsOpen) comport.Close();
            //else
            //{
            //    //�]�w��
            //    comport.BaudRate = 9600;
            //    comport.DataBits = 8;
            //    comport.StopBits = StopBits.One;
            //    comport.Parity = Parity.None;
            //    comport.PortName = "COM1";
            //    try
            //    {
            //        comport.Open();
            //    }
            //    catch 
            //    {
            //        //MessageBox.Show(ex.Message);
            //       // return;
            //    }
            //}

            if (comport.IsOpen)
            {
                MessageBox.Show("�o�����w���\�s��");
                btnPrintTest.Enabled = true;
             
            }

            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            comboBox3.Text = "�P�f���";
            //comboBox4.Text = "�ɧ�";

            //comboBox2.Items.Clear();

            //System.Data.DataTable dt3 = GetOrderData3V();

            //for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            //{
            //    comboBox2.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            //}

            //comboBox2.Items.Add("����");
        }
        public static void Order(SerialPort printer, byte[] command)
        {
            printer.Write(command, 0, command.Length);
        }
        private void btnPrintTest_Click(object sender, EventArgs e)
        {
            //PRINT
            int f1 = 0;
            int f2 = 0;
            for (int j = 0; j <= dataGridView1.SelectedRows.Count - 1; j++)
            {
                string F = dataGridView1.SelectedRows[j].Cells["�o�����X"].Value.ToString();
                string PRINT = dataGridView1.SelectedRows[j].Cells["PRINT"].Value.ToString();
                if (String.IsNullOrEmpty(F))
                {
                  
                    f1 = 1;
                }
                if (PRINT == "True")
                {
                    f2 = 1;
                }
            }

            if (f1 == 1)
            {
                MessageBox.Show("�z���ťժ��o�����X");
                return;
            }
            if (f2 == 1)
            {
      
                DialogResult result2;
                result2 = MessageBox.Show("�z���o���w�C�L�L�A�T�w�O�_�n���ƦC�L", "YES/NO", MessageBoxButtons.YesNo);
                if (result2 == DialogResult.No)
                {
                    return;
                }
            }

            if (dataGridView1.Rows.Count == 0)
            {
                return;
            }

            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("�п�ܦC�L���C");
                return;
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                for (int j = dataGridView1.SelectedRows.Count - 1; j >= 0; j--)
                {
                    string �o�����X = dataGridView1.SelectedRows[j].Cells["�o�����X"].Value.ToString();


                    sb.Append(�o�����X + " / ");
                }
                sb.Remove(sb.Length - 2, 2);
                MESS =sb.ToString();
            }


                    DialogResult result;
                    result = MessageBox.Show("�нT�w�O�_�n�C�L�o�����X " + MESS, "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {


                for (int j = dataGridView1.SelectedRows.Count - 1; j >= 0; j--)
                {


                    string F2 = dataGridView1.SelectedRows[j].Cells["ID"].Value.ToString();
                    if (F2.IndexOf("/") != -1)
                    {
                        F2 = F2.Substring(0, F2.IndexOf("/")) + "~" + F2.Substring(F2.LastIndexOf("/") + 1, 10);
                    }
                    string FF2 = dataGridView1.SelectedRows[j].Cells["ID2"].Value.ToString();
                    string INV = dataGridView1.SelectedRows[j].Cells["�νs"].Value.ToString();
                    string DOC = dataGridView1.SelectedRows[j].Cells["�P�f���"].Value.ToString();
                    string ���|���B = dataGridView1.SelectedRows[j].Cells["���|���B"].Value.ToString();
                    string �Ȥ�W�� = dataGridView1.SelectedRows[j].Cells["�Ȥ�W��"].Value.ToString();
                    string �ҵ|���O = dataGridView1.SelectedRows[j].Cells["�ҵ|���O"].Value.ToString();
                    string �o�����X = dataGridView1.SelectedRows[j].Cells["�o�����X"].Value.ToString();

                    int MAN = �Ȥ�W��.IndexOf("�֪�");
                    int MAN2 = �Ȥ�W��.IndexOf("�E���t�a");
                    int MAN3 = �Ȥ�W��.IndexOf("�ڷ�");
                    bool openMoneyBox_BeforePrinting = true;
                    bool openMoneyBox_AfterPrinting = true;

                    comport.Encoding = Encoding.Default;

                    // comport.Order(Command.ResetPrinter);    //��l�L���
                    comport.Write(Command.ResetPrinter, 0, Command.ResetPrinter.Length);
                  
                    comport.Write(Command.StubAndReceiver, 0, Command.StubAndReceiver.Length);

                    if (openMoneyBox_BeforePrinting)
                        //comport.Order(Command.OpenMoneyBox1);
                        comport.Write(Command.OpenMoneyBox1, 0, Command.OpenMoneyBox1.Length);
                    System.Data.DataTable T1 = GetOrderData31(FF2);
           
                    string DOCDATE = DOC.Substring(0, 4) + "/" + DOC.Substring(4, 2) + "/" + DOC.Substring(6, 2);
                    comport.WriteLine("���׹�~�ѥ��������q");
                    comport.WriteLine("��~�H�νs: 22468373");
                    comport.WriteLine("�x�_������Ϸs��G��");
                    comport.WriteLine("257��5�Ӥ�3 TEL:87922800");
                    comport.WriteLine("POS# ARMAS-001");

                    comport.WriteLine(DOCDATE);
                    if (!String.IsNullOrEmpty(INV))
                    {
                        comport.WriteLine("�Τ@�s��: " + INV);
                    }
                    comport.WriteLine("------------------------");

            
                    if (T1.Rows.Count > 0)
                    {
                        UpdateID2(FF2, "True", �o�����X);
                        if (T1.Rows.Count > 9 || ���|���B.Length > 4)
                        {
                            System.Data.DataTable TT1 = GetOrderData313(FF2);

                            string INVNAME = TT1.Rows[0]["INVNAME"].ToString();
                     
                            string QTY = "";
                            if (MAN != -1 || MAN2 != -1 || MAN3 != -1)
                            {
                             
                                QTY = TT1.Rows[0]["QTY1"].ToString();
                            }
                            else
                            {
                                QTY = TT1.Rows[0]["QTY"].ToString();
                            }

                            int QTYT = QTY.Length;
                            if (QTYT == 1)
                            {
                                QTY = "    " + QTY;
                            }
                            else if (QTYT == 2)
                            {
                                QTY = "   " + QTY;
                            }
                            else if (QTYT == 3)
                            {
                                QTY = "  " + QTY;
                            }
                            else if (QTYT == 4)
                            {
                                QTY = " " + QTY;
                            }


                            int TAMOUNT = Convert.ToInt32(TT1.Rows[0]["AMOUNT"].ToString());
                            string AMOUNT = TAMOUNT.ToString("#,##0");
                            int AMOUNTT = AMOUNT.Length;
                            if (AMOUNTT == 1)
                            {
                                AMOUNT = "       " + AMOUNT;
                            }
                            if (AMOUNTT == 2)
                            {
                                AMOUNT = "      " + AMOUNT;
                            }
                            if (AMOUNTT == 3)
                            {
                                AMOUNT = "     " + AMOUNT;
                            }
                            if (AMOUNTT == 4)
                            {
                                AMOUNT = "    " + AMOUNT;
                            }
                            if (AMOUNTT == 5)
                            {
                                AMOUNT = "   " + AMOUNT;
                            }
                            if (AMOUNTT == 6)
                            {
                                AMOUNT = "  " + AMOUNT;
                            }
                            if (AMOUNTT == 7)
                            {
                                AMOUNT = " " + AMOUNT;
                            }

                            comport.WriteLine(INVNAME + QTY + AMOUNT + �ҵ|���O);
                        }
                        else
                        {
                            for (int i = 0; i <= T1.Rows.Count - 1; i++)
                            {
                                string INVNAME = T1.Rows[i]["INVNAME"].ToString();
                 
                                string QTY = T1.Rows[i]["QTY"].ToString();
                                int QTYT = QTY.Length;
                                if (QTYT == 1)
                                {
                                    QTY = "     " + QTY;
                                }
                                else if (QTYT == 2)
                                {
                                    QTY = "    " + QTY;
                                }
                                else if (QTYT == 3)
                                {
                                    QTY = "   " + QTY;
                                }
                                else if (QTYT == 4)
                                {
                                    QTY = "  " + QTY;
                                }
                                else if (QTYT == 5)
                                {
                                    QTY = " " + QTY;
                                }



                                int TAMOUNT = Convert.ToInt32(T1.Rows[i]["AMOUNT"].ToString());
                                string AMOUNT = TAMOUNT.ToString("#,##0");
                                int AMOUNTT = AMOUNT.Length;
                                if (AMOUNTT == 5)
                                {
                                    AMOUNT = "  " + AMOUNT;
                                }
                                if (AMOUNTT == 6)
                                {
                                    AMOUNT = " " + AMOUNT;
                                }
                                if (AMOUNTT == 3)
                                {
                                    AMOUNT = "    " + AMOUNT;
                                }
                                if (AMOUNTT == 2)
                                {
                                    AMOUNT = "     " + AMOUNT;
                                }
                                if (AMOUNTT == 1)
                                {
                                    AMOUNT = "      " + AMOUNT;
                                }
                             
                                comport.WriteLine(INVNAME + QTY + AMOUNT + �ҵ|���O);
                                // }

                            }
                        }


                        int T���B = Convert.ToInt32(���|���B);
                        string ���B = T���B.ToString("#,##0");
                        string ���B2 = T���B.ToString("#,##0");
                        int ���BT = ���B.Length;
                        string �K�|���B = "";
                        if (���BT == 5)
                        {
                            ���B = "            " + ���B;
                        }
                        if (���BT == 6)
                        {
                            ���B = "           " + ���B;
                        }
                        if (���BT == 7)
                        {
                            ���B = "          " + ���B;
                        }
                        if (���BT == 3)
                        {
                            ���B = "              " + ���B;
                        }

                        if (���BT == 5)
                        {
                            �K�|���B = "          " + ���B2;
                        }
                        if (���BT == 6)
                        {
                            �K�|���B = "         " + ���B2;
                        }
                        if (���BT == 7)
                        {
                            �K�|���B = "        " + ���B2;
                        }
                        if (���BT == 3)
                        {
                            �K�|���B = "            " + ���B2;
                        }
       

                        int T�`�p = Convert.ToInt32(���|���B);
                        string �`�p = T�`�p.ToString("#,##0");
                        int �`�pT = �`�p.Length;
                        if (�`�pT == 5)
                        {
                            �`�p = "              " + �`�p;
                        }
                        if (�`�pT == 6)
                        {
                            �`�p = "             " + �`�p;
                        }
                        if (�`�pT == 7)
                        {
                            �`�p = "            " + �`�p;
                        }
                        if (�`�pT == 3)
                        {
                            �`�p = "                " + �`�p;
                        }

                        comport.WriteLine("------------------------");
                        comport.WriteLine("�p�p:" + ���B + �ҵ|���O);
                 
                        comport.WriteLine("========================");
                        comport.WriteLine("�`�p:" + �`�p);

                        int F = 0;
                        int F3 = 0;
                     

                        comport.Write(Command.MoveLines(1), 0, Command.MoveLines(1).Length);
                        comport.WriteLine("PO# " + F2);
                        comport.WriteLine("SO# " + FF2);
                  

                        //comport.WriteLine("from Acmepoint MIS");

                        // comport.Order(Command.MoveLines(20));   //���쩱���B
                        comport.Write(Command.MoveLines(20), 0, Command.MoveLines(20).Length);
                        // comport.Order(Command.PrintMark);       //�L����
                        comport.Write(Command.PrintMark, 0, Command.PrintMark.Length);
                        //  comport.Order(Command.NewPage);         //����
                        comport.Write(Command.NewPage, 0, Command.NewPage.Length);
                  
                        if (openMoneyBox_AfterPrinting)
                            // comport.Order(Command.OpenMoneyBox1);
                            comport.Write(Command.OpenMoneyBox1, 0, Command.OpenMoneyBox1.Length);
                           
                    }
                }
                EXEC();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            EXEC();
        }

        private void TOTAL2GG(System.Data.DataTable dt)
        {
            dtGetAcmeStageG = MakeTableCombineGG();

            System.Data.DataTable DT1 = dt;
            DataRow dr = null;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                dr = dtGetAcmeStageG.NewRow();
                string ID = DT1.Rows[i]["ID"].ToString().Trim();
                string ID2 = DT1.Rows[i]["ID2"].ToString().Trim();
   string CUSTID = DT1.Rows[i]["�Ȥ�s��"].ToString().Trim();
                string CUSTNAME = DT1.Rows[i]["�Ȥ�W��"].ToString().Trim();
                dr["ID2"] = ID2;
                dr["�Ȥ�s��"] = CUSTID;
                dr["�Ȥ�W��"] = CUSTNAME;
                decimal ���|���B= Convert.ToDecimal(DT1.Rows[i]["���|���B"]);
                dr["���|���B"] = ���|���B;
                dr["�|�B"] = Convert.ToDecimal(DT1.Rows[i]["�|�B"]);
                dr["�K�|���B"] = Convert.ToDecimal(DT1.Rows[i]["�K�|���B"]);
                dr["����覡"] = DT1.Rows[i]["����覡"].ToString().Trim();
                dr["�P�f���"] = DT1.Rows[i]["�P�f���"].ToString().Trim();
                dr["�~��"] = DT1.Rows[i]["�~��"].ToString().Trim();
                string DEP = DT1.Rows[i]["DEPTID"].ToString().Trim();
                dr["�b���k��"] = DT1.Rows[i]["�b���k��"].ToString().Trim();
                dr["����"] = DEP;
                string INVOTYPE = DT1.Rows[i]["�o�����O"].ToString().Trim();
                dr["�o�����O"] = INVOTYPE;
                System.Data.DataTable L1 = GetREMARK(ID2,comboBox8.Text);
                if (L1.Rows.Count > 0)
                {
                    string remark = Convert.ToString(L1.Rows[0][0].ToString());

                    string[] sArray = remark.Split('\r');
                    int F2 = 0;
                    foreach (string F in sArray)
                    {
                        F2++;
                    }
                    if (F2 > 2)
                    {
                        string tmpOrder = sArray[2];
                        int FS = tmpOrder.IndexOf(":");
                        if (FS != -1)
                        {
                            string[] sArray1 = tmpOrder.Split(':');
                            string H1 = sArray1[1];

                            if (!String.IsNullOrEmpty(H1))
                            {
                                System.Data.DataTable T1 = GetCARD(H1);
                                if (T1.Rows.Count > 0)
                                {
                                    dr["�d�����|�X"] = T1.Rows[0][0].ToString();
                                }

                            }
                        }


                        if (string.IsNullOrEmpty(dr["�d�����|�X"].ToString()))
                        {
                            if (F2 > 9)
                            {
                                string tmpOrder2 = sArray[9];
                                int INT1 = tmpOrder2.IndexOf("�d�����|�X");
                                if (INT1 != -1)
                                {
                                    string[] sArray12 = tmpOrder2.Split(':');
                                    string H2 = sArray12[1];

                                    dr["�d�����|�X"] = H2.ToString().Trim();
                                }

                            }
                        }
                    }
                }
                string CUSTNO = DT1.Rows[i]["�νs"].ToString().Trim();
                if (String.IsNullOrEmpty(CUSTNO))
                {
                    System.Data.DataTable GG1 = GetCUSTNO(ID2);
                    if (GG1.Rows.Count > 0)
                    {
                        CUSTNO = GG1.Rows[0][0].ToString();
                    }
      
                }
                if (String.IsNullOrEmpty(CUSTNO))
                {

                    if (L1.Rows.Count > 0)
                    {
                        string remark = Convert.ToString(L1.Rows[0][0].ToString());

                        int S1 = remark.IndexOf("7.�νs:");
                        if (S1 != -1)
                        {

                            //8.�o���a�}
                            CUSTNO = remark.Substring(S1 + 5, 9).Replace("8.�o���a�}", "").Replace("8.�o���a", "").Replace(":", "");

                        }
                    }
                }
                //if (INVOTYPE == "35")
                //{
                //    if (String.IsNullOrEmpty(CUSTNO))
                //    {
                //        dr["�|�B"] = 0;
                //    }
                //    else
                //    {
                //        dr["�|�B"] = Convert.ToDecimal(DT1.Rows[i]["�|�B"]);
                //    }
                //}

                dr["�|�B"] = Convert.ToDecimal(DT1.Rows[i]["�|�B"]);
                dr["�o�����X"] = DT1.Rows[i]["�o�����X"].ToString().Trim();
                dr["�o�����"] = DT1.Rows[i]["�o�����"].ToString().Trim();
                StringBuilder sb3 = new StringBuilder();
                System.Data.DataTable G3 = GetPAY(ID2);
                if (G3.Rows.Count > 0)
                {
                    for (int s = 0; s <= G3.Rows.Count - 1; s++)
                    {

                        DataRow dd = G3.Rows[s];
                        string FNO = dd["FNO"].ToString();
                        sb3.Append(FNO + "/");

                    }

                    if (sb3.Length > 0)
                    {
                        sb3.Remove(sb3.Length - 1, 1);
                    }
                    dr["���ڳ渹"] = sb3.ToString();
                }
      
                dr["���ڪ��B"] = DT1.Rows[i]["���ڪ��B"].ToString().Trim();
                string PRINT = DT1.Rows[i]["PRINT"].ToString().Trim();
                if (DEP == "C2")
                {
                    PRINT = "True";
                }
                dr["PRINT"] = PRINT;
                dr["PRINT2"] = DT1.Rows[i]["PRINT2"].ToString().Trim();
                dr["�������u"] = DT1.Rows[i]["�������u"].ToString().Trim();
                dr["�ҵ|���O"] = DT1.Rows[i]["�ҵ|���O"].ToString().Trim();
        
                StringBuilder sb2 = new StringBuilder();
                System.Data.DataTable G2 = GetOrderData4(ID2);
                System.Data.DataTable G1 = GetOrderData3TT(ID2);
                if (G1.Rows.Count > 0)
                {

                    for (int s = 0; s <= G1.Rows.Count - 1; s++)
                    {

                        DataRow dd = G1.Rows[s];
                        string BILLNO = dd["BILLNO"].ToString();
                        sb2.Append(BILLNO + "/");

                        System.Data.DataTable GZEN = GetZENBEN(BILLNO);
                        if (GZEN.Rows.Count > 0)
                        {
                            dr["�����ȥ�"] = "Y";
                        }
                        System.Data.DataTable GTON = GetTONBEN(BILLNO);
                        if (String.IsNullOrEmpty(CUSTNO))
                        {
                            if (GTON.Rows.Count > 0)
                            {
                                CUSTNO = GTON.Rows[0][0].ToString();
                            }
                        }

          
                        if (String.IsNullOrEmpty(CUSTNO))
                        {
                            System.Data.DataTable GTON2 = GetTONBEN2(BILLNO);
                            if (GTON2.Rows.Count > 0)
                            {
                                CUSTNO = GTON2.Rows[0][0].ToString();
                            }
                        }
                    }
                    if (sb2.Length > 0)
                    {
                        sb2.Remove(sb2.Length - 1, 1);
                    }
                    dr["ID"] = sb2.ToString();
                }
                if (ID == "4")
                {
                    dr["ID"] = "���J��";
                }
                dr["�νs"] = CUSTNO;
                UpdateCUSTINV(CUSTNO, ID2);
                string MONTH = DateTime.Now.ToString("yyyyMM");
                System.Data.DataTable SH1 = GETAR1(ID2);

                if (DEP == "C1")
                {
                    if (���|���B == 0)
                    {
                        if (SH1.Rows.Count == 0)
                        {
                            string INVNO = DateTime.Now.ToString("yyyyMMddHHmmss");
                            INSERTINVOICE("2", INVNO, "", GetMenu.Day(), "2", ID2, "500", CUSTID, MONTH, "36", "0", 0, "", 0, 0, CUSTNAME, CUSTNO, "", 1, 1, "", MONTH, 0, MONTH, "0", "True", "0", "�s�W", "0", "", "", "0", "", "0", "0", "0", "1", "A34", "Sharon", "");
                            UPDATEINVOICE(INVNO, ID2);
                         //   MessageBox.Show("�s�W���B0�o�� " + ID2);
                        }
                    }
                }

                //INSERTINVOICE("2", "JC40888720", "JC40888720", "20210118", "2", "202101180002", FLAG, CUST, MONTH, "35", "0", TAXTYPE, "", 10466, 524, CUSTNAME, "69773557", "", 1, INCLUDETAX, "", MONTH, 10990, MONTH, "0", "True", "0", "�s�W", "0", CUSTADD, "", "0", "", "0", "0", "0", "1", "A34", "Sharon", "");
                //  }
                if (ID == "2" )
                {
                    if (String.IsNullOrEmpty(dr["ID"].ToString()))
                    {
                        if (G2.Rows.Count == 0)
                        {
                            dtGetAcmeStageG.Rows.Add(dr);
                        }
                    }
                }
                else
                {
                    dtGetAcmeStageG.Rows.Add(dr);
                }
            }


        }
        public void UPDATEINVOICE(string InvoiceNO, string SrcBillNO)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
       
            sb.Append("UPDATE COMBILLACCOUNTS SET INVOBILLNO=@InvoiceNO,InvoFlag=2 WHERE  FundBillNo  =@SrcBillNO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@InvoiceNO", InvoiceNO));
            command.Parameters.Add(new SqlParameter("@SrcBillNO", SrcBillNO));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        private System.Data.DataTable GETAR1(string BillNO)
        {


            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT FLAG FROM COMINVOICE  WHERE SRCBILLNO=@BillNO         ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable MakeTableCombineGG()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("ID2", typeof(string));
            dt.Columns.Add("����", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("�K�|���B", typeof(decimal));
            dt.Columns.Add("�|�B", typeof(decimal));
            dt.Columns.Add("���|���B", typeof(decimal));
            dt.Columns.Add("����覡", typeof(string));
            dt.Columns.Add("�d�����|�X", typeof(string));
            dt.Columns.Add("�P�f���", typeof(string));
            dt.Columns.Add("�νs", typeof(string));
            dt.Columns.Add("�~��", typeof(string));
            dt.Columns.Add("�b���k��", typeof(string));
            dt.Columns.Add("�o�����O", typeof(string));
            dt.Columns.Add("�o�����X", typeof(string));
            dt.Columns.Add("�o�����", typeof(string));
            dt.Columns.Add("���ڳ渹", typeof(string));
            dt.Columns.Add("���ڪ��B", typeof(string));
            dt.Columns.Add("PRINT", typeof(string));
            dt.Columns.Add("PRINT2", typeof(string));
            dt.Columns.Add("�������u", typeof(string));
            dt.Columns.Add("�ҵ|���O", typeof(string));
            dt.Columns.Add("�����ȥ�", typeof(string));
         
            return dt;
        }

        private void EXEC()
        {

            System.Data.DataTable dt = null;

            if (comboBox8.Text !="")
            {
                dt = GetOrderData4F(comboBox8.Text);
            }
            else
            {
                dt = GetOrderData3("2");
            }
            TOTAL2GG(dt);
            if (dt.Rows.Count == 0)
            {
                System.Data.DataTable dt2 = GetOrderData3("1");
                dataGridView1.DataSource = dt2;

            }
            else
            {

                dataGridView1.DataSource = dtGetAcmeStageG;
                dataGridView1.Columns["�o�����X"].ReadOnly = false;

                DataRow row;
                //�[�J�@���X�p
                Int32[] Total = new Int32[dtGetAcmeStageG.Columns.Count - 1];

                for (int i = 0; i <= dtGetAcmeStageG.Rows.Count - 1; i++)
                {

                    for (int j = 4; j <= 6 ; j++)
                    {
                        try
                        {
                            Total[j - 1] += Convert.ToInt32(dtGetAcmeStageG.Rows[i][j]);
                        }
                        catch
                        {
                            Total[j - 1] += 0;
                        }

                    }
                }
                row = dtGetAcmeStageG.NewRow();

                row[2] = "�X�p";
                for (int j = 4; j <= 6; j++)
                {
                    row[j] = Total[j - 1];

        
                }
                dtGetAcmeStageG.Rows.Add(row);



            }
        }
        private System.Data.DataTable GetOrderData3(string A)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("                           SELECT DISTINCT  '1' ID,O.BillNO ID2,   ");
            sb.Append("                                                                                                     U.ID �Ȥ�s��,   U.FullName  �Ȥ�W��,ISNULL(S.TOTAL,0) �K�|���B,CAST(ISNULL(S.TAX,0) AS INT) �|�B,CAST(ISNULL(S.TOTAL+S.TAX,0) AS INT) ���|���B,       ");
            sb.Append("                                                CASE T1.GatherStyle WHEN 0 THEN '�f��' WHEN 1 THEN '����' WHEN 2 THEN '�뵲' WHEN 3 THEN T1.GatherOther END ����覡  ");
            sb.Append("                                                     ,O.BILLDATE �P�f���,U.TaxNo  �νs,I.InvoiceNO �o�����X,I.InvoiceDate �o�����,    ");
            sb.Append("                                                 (Select Convert(int, SUM(H.Total))    ");
            sb.Append("                                                 From ComFundSub H    ");
            sb.Append("                                                 Where S.Flag = H.OriginFlag     ");
            sb.Append("                                                 And S.FundBillNO = H.OriginBillNO    ");
            sb.Append("                                                 And H.OriginFlag <> 0    ");
            sb.Append("                                                 And Left(H.OriginBillNO,1) <> '*') as ���ڪ��B,   ");
           // sb.Append("                                                 case isnull(I.PrintMan,'False') when '' then 'False' else isnull(I.PrintMan,'False') end [PRINT],");
            sb.Append("                                               case isnull(I.PRINTER,'') when '' then 'False' else 'True' end [PRINT],");
            sb.Append("                          I.PRINTER [PRINT2],");
            sb.Append("  CASE WHEN U.FullName LIKE '%-�i����%' THEN U.FullName  WHEN U.FullName LIKE '%����%' THEN U.FullName  WHEN U.FullName LIKE '%����%'  THEN U.FullName WHEN U.FullName LIKE '%�෽�A��%' THEN U.FullName  END �������u");
            sb.Append("                                                      ,CASE S.TaxType WHEN 0 THEN 'TX' WHEN 1 THEN 'NX' END �ҵ|���O,I.InvoiceType �o�����O,P.PersonName �~��,Q.FullName �b���k��,S.DEPTID    FROM ComProdRec O  ");
            sb.Append("                                           INNER join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)  ");
            sb.Append("                                            LEFT join comCustomer U On  U.ID=S.CustID    ");
            sb.Append("                                            INNER Join comProduct B On B.ProdID =o.ProdID   ");
            sb.Append("                                                LEFT Join comInvoice I On  O.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1  ");
            sb.Append("                                                    INNER Join StkBillSUB T0 On  O.BillNO=T0.BillNO AND  O.RowNO =T0.RowNO ");
            sb.Append("                                                       INNER Join StkBillMAIN T1 On T0.BillNO=T1.BillNO  ");
            sb.Append(" left join comPerson P ON (S.Salesman=P.PersonID)  						 Left join comCustomer Q ON S.DueTo = Q.ID And Q.Flag =S.CustFlag       ");
            sb.Append("                                       WHERE O.Flag =500   ");
            sb.Append("  AND  O.BILLDATE BETWEEN @CreateDate AND @CreateDate1 ");     
            if (A == "1")
            {
                sb.Append(" AND 1 = 2 ");
            }

            if (checkBox4.Checked)
            {
                sb.Append(" and  S.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  S.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (comboBox1.Text != "")
            {
                sb.Append("  AND CASE S.TaxType WHEN 0 THEN 'TX' WHEN 1 THEN 'NX' END   = @TX ");
            }
            if (comboBox2.Text != "")
            {
                sb.Append("  AND CASE T1.GatherStyle WHEN 0 THEN '�f��' WHEN 1 THEN '����' WHEN 2 THEN '�뵲' WHEN 3 THEN T1.GatherOther END   = @TTYPE ");
            }
            if (textBox5.Text != "" )
            {
                sb.Append(" and  O.BillNO=@BillNO ");
            }
            if (comboBox5.Text != "")
            {
                sb.Append(" and I.InvoiceType=@InvoiceType ");
            }
            if (textBox4.Text != "")
            {
                sb.Append(" and I.InvoiceNO=@InvoiceNO ");
            }
            if (textBox6.Text != "" && textBox9.Text != "")
            {
                sb.Append(" and  I.InvoiceDate between @InvoiceDate1 and @InvoiceDate2 ");
            }
            if (comboBox6.Text != "")
            {
                if (comboBox6.Text == "�w�}��")
                {
                    sb.Append(" and isnull(I.PrintMan,'False')='True' ");
                }
                if (comboBox6.Text == "���}��")
                {
                    sb.Append("  and case isnull(I.PrintMan,'False') when '' then 'False' else isnull(I.PrintMan,'False') end='False' ");
                }
                //if (comboBox6.Text == "�w�}��")
                //{
                //    sb.Append(" and isnull(I.PRINTER,'') <> '' ");
                //}
                //if (comboBox6.Text == "���}��")
                //{
                //    sb.Append(" and isnull(I.PRINTER,'') = '' ");
                //}
            }
            if (comboBox7.Text != "")
            {
                sb.Append(" and  I.PRINTER=@PRINTER ");
            }
            if (textBox10.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  S.DEPTID BETWEEN @C1 AND @C2 ");
            }
            //,S.DEPTID 
            sb.Append("           UNION ALL");
            sb.Append("                                               SELECT DISTINCT  '1' ID,I.SrcBillNO ID2,                              ");
            sb.Append("                                                                                                      U.ID �Ȥ�s��,   U.FullName  �Ȥ�W��,Amount  �K�|���B,TaxAmt  �|�B,Amount+TaxAmt ���|���B,      ");
            sb.Append("                                                                                        ''����覡    ");
            sb.Append("                                                                                             ,I.InvoiceDate  �P�f���,TaxRegNO  �νs,I.InvoiceNO �o�����X,I.InvoiceDate �o�����,       ");
            sb.Append("                                                                                            ''���ڪ��B,     ");
            sb.Append("                                               case isnull(I.PRINTER,'') when '' then 'False' else 'True' end [PRINT],");
            sb.Append("                                                                           case isnull(I.PRINTER,'False') when '' then 'False' else isnull(I.PRINTER,'False') end [PRINT2], ");
            sb.Append("                                       '' �������u                ,'FX' �ҵ|���O,I.InvoiceType �o�����O,P.PersonName �~��,Q.FullName �b���k��,S.DeptID DEPTID     ");
            sb.Append("        FROM comInvoice I");
            sb.Append("																      left join COMBILLACCOUNTS S ON (I.SrcBillNO =S.FundBillNo AND S.Flag =600)    ");
            sb.Append("                                            LEFT join comCustomer U On  U.ID=S.CustID    ");
            sb.Append(" left join comPerson P ON (S.Salesman=P.PersonID)  						 Left join comCustomer Q ON S.DueTo = Q.ID And Q.Flag =S.CustFlag       ");
            sb.Append("																   WHERE IsCancel =1  ");
            sb.Append("                                    AND  I.InvoiceDate BETWEEN @CreateDate AND @CreateDate1   ");

            if (comboBox6.Text != "")
            {
                if (comboBox6.Text == "�w�}��")
                {
                    sb.Append(" and isnull(I.PrintMan,'False')='True' ");
                }
                if (comboBox6.Text == "���}��")
                {
                    sb.Append("  and case isnull(I.PrintMan,'False') when '' then 'False' else isnull(I.PrintMan,'False') end='False' ");
                }
            }
            //if (comboBox6.Text == "�w�}��")
            //{
            //    sb.Append(" and isnull(I.PRINTER,'') <> '' ");
            //}
            //if (comboBox6.Text == "���}��")
            //{
            //    sb.Append(" and isnull(I.PRINTER,'') = '' ");
            //}
            if (checkBox4.Checked)
            {
                sb.Append(" and  S.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  S.[CustID] between @CustID1 and @CustID2 ");
                }
            }


            if (textBox10.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  S.DEPTID BETWEEN @C1 AND @C2 ");
            }
            if (A == "1" || textBox5.Text != "")
            {
                sb.Append(" AND 1 = 2 ");
            }

            if (checkBox1.Checked)
            {
                sb.Append("           UNION ALL");
                sb.Append("                         SELECT DISTINCT  '2' ID,O.BillNO ID2, ");
                sb.Append("                                                                                                     U.ID �Ȥ�s��,   U.FullName  �Ȥ�W��,ISNULL(I.AMOUNT,0) �K�|���B,CAST(ISNULL(I.TaxAmt,0) AS INT) �|�B,CAST(ISNULL(I.AMOUNT+I.TaxAmt,0) AS INT) ���|���B,       ");
                sb.Append("                                                    CASE T1.GatherStyle WHEN 0 THEN '�f��' WHEN 1 THEN '����' WHEN 2 THEN '�뵲' WHEN 3 THEN T1.GatherOther END ����覡  ");
                sb.Append("                                                         ,O.BILLDATE �P�f���,CASE WHEN U.ClassID IN (011,013,015,014,018,019,020,021,022,026,027,028,029) THEN U.TaxNo ELSE I.TaxRegNO END �νs,'' �o�����X,''�o�����,   ");
                sb.Append("                                                     (Select Convert(int, SUM(H.Total))    ");
                sb.Append("                                                     From ComFundSub H    ");
                sb.Append("                                                     Where S.Flag = H.OriginFlag     ");
                sb.Append("                                                     And S.FundBillNO = H.OriginBillNO    ");
                sb.Append("                                                     And H.OriginFlag <> 0    ");
                sb.Append("                                                     And Left(H.OriginBillNO,1) <> '*') as ���ڪ��B,   ");
                sb.Append("                                                         case isnull(I.PrintMan,'False') when '' then 'False' else isnull(I.PrintMan,'False') end [PRINT],");
                sb.Append("                                                         '' [PRINT2],");
                sb.Append("   CASE WHEN U.FullName LIKE '%-�i����%' THEN U.FullName  WHEN U.FullName LIKE '%����%' THEN U.FullName  WHEN U.FullName LIKE '%����%' THEN U.FullName  WHEN U.FullName LIKE '%�෽�A��%' THEN U.FullName   END �������u,CASE S.TaxType WHEN 0 THEN 'TX' WHEN 1 THEN 'NX' END �ҵ|���O,I.InvoiceType �o�����O,P.PersonName �~��,Q.FullName �b���k�� ,S.DEPTID      FROM ComProdRec O    ");
                sb.Append("                                                             left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)   ");
                sb.Append("                                                              left join comCustomer U On  U.ID=S.CustID  AND U.Flag =1  ");
                sb.Append("                                            Left Join comProduct B On B.ProdID =o.ProdID   ");
                sb.Append("                                                                  Left Join comInvoice I On  O.BillNO=I.SrcBillNO AND I.Flag =2  AND I.IsCancel <> 1  ");
                sb.Append("                                                                      Left Join StkBillSUB T0 On  O.BillNO=T0.BillNO AND  O.RowNO =T0.RowNO  ");
                sb.Append("                                                                         Left Join StkBillMAIN T1 On T0.BillNO=T1.BillNO   ");
                sb.Append(" left join comPerson P ON (S.Salesman=P.PersonID)  	 Left join comCustomer Q ON S.DueTo = Q.ID And Q.Flag =S.CustFlag       ");
                sb.Append("                                                         WHERE ISNULL(I.InvoiceNO,'') <> '' AND I.SrcSysNO=2  and I.InvoiceType <> 36  ");
                sb.Append("                                    AND  O.BILLDATE BETWEEN @CreateDate AND @CreateDate1   ");
                if (A == "1")
                {
                    sb.Append(" AND 1 <> 2 ");
                }

                if (checkBox4.Checked)
                {
                    sb.Append(" and  S.[CustID] in ( " + c + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  S.[CustID] between @CustID1 and @CustID2 ");
                    }
                }
                if (comboBox1.Text != "")
                {
                    sb.Append("  AND CASE S.TaxType WHEN 0 THEN 'TX' WHEN 1 THEN 'NX' END   = @TX ");
                }
                if (comboBox2.Text != "")
                {
                    sb.Append("  AND CASE T1.GatherStyle WHEN 0 THEN '�f��' WHEN 1 THEN '����' WHEN 2 THEN '�뵲' WHEN 3 THEN T1.GatherOther END   = @TTYPE ");
                }
                if (textBox5.Text != "")
                {
                    sb.Append(" and  O.BillNO=@BillNO ");
                }
                if (comboBox5.Text != "")
                {
                    sb.Append(" and I.InvoiceType=@InvoiceType ");
                }
                if (textBox4.Text != "")
                {
                    sb.Append(" and I.InvoiceNO=@InvoiceNO ");
                }
                if (textBox6.Text != "" && textBox9.Text != "")
                {
                    sb.Append(" and  I.InvoiceDate between @InvoiceDate1 and @InvoiceDate2 ");
                }
                //if (comboBox6.Text != "")
                //{
                //    if (comboBox6.Text == "�w�}��")
                //    {
                //        sb.Append(" and isnull(I.PrintMan,'False')='True' ");
                //    }
                //    if (comboBox6.Text == "���}��")
                //    {
                //        sb.Append("  and case isnull(I.PrintMan,'False') when '' then 'False' else isnull(I.PrintMan,'False') end='False' ");
                //    }
                //}
                if (comboBox6.Text == "�w�}��")
                {
                    sb.Append(" and isnull(I.PRINTER,'') <> '' ");
                }
                if (comboBox6.Text == "���}��")
                {
                    sb.Append(" and isnull(I.PRINTER,'') = '' ");
                }
                if (comboBox7.Text != "")
                {
                    sb.Append(" and  I.PRINTER=@PRINTER ");
                }
                if (textBox10.Text != "" && textBox12.Text != "")
                {
                    sb.Append(" and  S.DEPTID BETWEEN @C1 AND @C2 ");
                }
                sb.Append("  UNION ALL      ");
                sb.Append(" SELECT DISTINCT  '4' ID,T0.FundBillNO ID2,  ");
                sb.Append(" U.ID �Ȥ�s��,   U.FullName  �Ȥ�W��,ISNULL(T0.TOTAL,0) �K�|���B,CAST(ISNULL(T0.TAX,0) AS INT) �|�B,CAST(ISNULL(T0.TOTAL+T0.TAX,0) AS INT) ���|���B,     ");
                sb.Append(" CASE T1.GatherStyle WHEN 0 THEN '�f��' WHEN 1 THEN '����' WHEN 2 THEN '�뵲' WHEN 3 THEN T1.GatherOther END ����覡  ");
                sb.Append(" ,T0.BILLDATE  �P�f���,U.TaxNo   �νs,I.InvoiceNO �o�����X,I.InvoiceDate �o�����,         ");
                sb.Append(" V.Offset ���ڪ��B,       ");
                sb.Append(" case isnull(I.PrintMan,'False') when '' then 'False' else isnull(I.PrintMan,'False') end [PRINT],    ");
                sb.Append(" case isnull(I.PRINTER,'False') when '' then 'False' else isnull(I.PRINTER,'False') end [PRINT2],   ");
                sb.Append(" '' �������u                ,CASE I.TaxType WHEN 0 THEN 'TX' WHEN 1 THEN 'NX' END �ҵ|���O,I.InvoiceType �o�����O,P.PersonName �~��,Q.FullName �b���k��,T0.DEPTID     ");
                sb.Append(" FROM comBillAccounts T0    ");
                sb.Append(" INNER JOIN comCostMain T1 ON (T1.CostBillNo =T0.FundBillNO)  ");
                sb.Append(" LEFT join comCustomer U On  U.ID=T0.CustID AND U.Flag =1");
                sb.Append(" LEFT JOIN comInvoice I ON (I.SrcBillNO =T0.FundBillNO)  ");
                sb.Append(" LEFT JOIN (Select  A.Offset,A.OriginBillNO From ComFundSub A  Inner Join ComFundMain B On A.Flag = B.Flag And A.FundBillNO = B.FundBillID  Where B.HasCheck = 1 And A.OriginFlag = 595 ) V ");
                sb.Append(" ON (V.OriginBillNO = I.SrcBillNO) ");
                sb.Append(" left join comPerson P ON (T0.Salesman=P.PersonID)     	 	 Left join comCustomer Q ON T0.DueTo = Q.ID And Q.Flag =T0.CustFlag         ");
                sb.Append(" WHERE  T0.Flag =595 ");
                sb.Append(" AND  T0.BILLDATE BETWEEN @CreateDate AND @CreateDate1      ");
                if (A == "1")
                {
                    sb.Append(" AND 1 = 2 ");
                }
       
            }
            if (comboBox3.Text == "�q�ʳ渹")
            {
                sb.Append(" ORDER BY ID");
            }
            else if (comboBox3.Text == "�P�f�渹")
            {
                sb.Append(" ORDER BY ID2");
            }

            else if (comboBox3.Text == "�P�f���")
            {
                sb.Append(" ORDER BY �P�f���");
            }
            else if (comboBox3.Text == "�o�����X")
            {
                sb.Append(" ORDER BY �o�����X");
            }
     


            if (comboBox4.Text == "�ɧ�")
            {
                sb.Append(" ASC");
            }
            else if (comboBox4.Text == "����")
            {
                sb.Append(" DESC");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 3600;
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CreateDate1", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@TX", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@TTYPE", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@BillNO", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceType", comboBox5.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceNO", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceDate1", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceDate2", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@PRINTER", comboBox7.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "SALES");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData4F(string DOCTYPE)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("                           SELECT DISTINCT  '1' ID,O.BillNO ID2,   ");
            sb.Append("                                                                                                     U.ID �Ȥ�s��,   U.FullName  �Ȥ�W��,CAST(ISNULL(I.AMOUNT,0) AS INT) �K�|���B ,CAST(ISNULL(I.TaxAmt,0) AS INT) �|�B,ISNULL(I.AMOUNT+I.TaxAmt,0) ���|���B,       ");
            sb.Append("                                                CASE T1.GatherStyle WHEN 0 THEN '�f��' WHEN 1 THEN '����' WHEN 2 THEN '�뵲' WHEN 3 THEN T1.GatherOther END ����覡  ");
            sb.Append("                                                     ,O.BILLDATE �P�f���,CASE WHEN U.ClassID IN (011,013,015,014,018,019,020,021,022,026,027,028,029) THEN U.TaxNo ELSE I.TaxRegNO END �νs,I.InvoiceNO �o�����X,I.InvoiceDate �o�����,    ");
            sb.Append("                                                 (Select Convert(int, SUM(H.Total))    ");
            sb.Append("                                                 From ComFundSub H    ");
            sb.Append("                                                 Where S.Flag = H.OriginFlag     ");
            sb.Append("                                                 And S.FundBillNO = H.OriginBillNO    ");
            sb.Append("                                                 And H.OriginFlag <> 0    ");
            sb.Append("                                                 And Left(H.OriginBillNO,1) <> '*') as ���ڪ��B,   ");
            sb.Append("                                                 case isnull(I.PrintMan,'False') when '' then 'False' else isnull(I.PrintMan,'False') end [PRINT],");
            sb.Append("                          I.PRINTER [PRINT2],");
            sb.Append("  CASE WHEN U.FullName LIKE '%-�i����%' THEN U.FullName  WHEN U.FullName LIKE '%����%' THEN U.FullName  WHEN U.FullName LIKE '%����%'  THEN U.FullName WHEN U.FullName LIKE '%�෽�A��%' THEN U.FullName  END �������u");
            sb.Append("                                                      ,CASE S.TaxType WHEN 0 THEN 'TX' WHEN 1 THEN 'NX' END �ҵ|���O,I.InvoiceType �o�����O ,P.PersonName �~��,Q.FullName �b���k��,S.DEPTID  FROM ComProdRec O  ");
            sb.Append("                                           LEFT join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo  AND  CASE O.Flag WHEN 701 THEN 698 ELSE O.Flag END=S.Flag)  ");
            sb.Append("                                            LEFT join comCustomer U On  U.ID=S.CustID  AND U.Flag =1  ");
            sb.Append("                                            LEFT Join comProduct B On B.ProdID =o.ProdID   ");
            sb.Append("                                                LEFT Join comInvoice I On  O.BillNO=I.SrcBillNO AND I.Flag =4 AND I.IsCancel <> 1  ");
            sb.Append("                                                    LEFT Join StkBillSUB T0 On  O.BillNO=T0.BillNO AND  O.RowNO =T0.RowNO ");
            sb.Append("                                                       LEFT Join StkBillMAIN T1 On T0.BillNO=T1.BillNO  ");
            sb.Append(" left join comPerson P ON (S.Salesman=P.PersonID)  						 Left join comCustomer Q ON S.DueTo = Q.ID And Q.Flag =S.CustFlag       ");
            if (DOCTYPE == "�P�h")
            {
                sb.Append("                                       WHERE O.Flag = 600 ");
            }
            if (DOCTYPE == "����")
            {
                sb.Append("                                       WHERE O.Flag  =  701  ");
            }
            sb.Append("  AND  O.BILLDATE BETWEEN @CreateDate AND @CreateDate1 ");

            if (checkBox4.Checked)
            {
                sb.Append(" and  S.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  S.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (comboBox1.Text != "")
            {
                sb.Append("  AND CASE S.TaxType WHEN 0 THEN 'TX' WHEN 1 THEN 'NX' END   = @TX ");
            }
            if (comboBox2.Text != "")
            {
                sb.Append("  AND CASE T1.GatherStyle WHEN 0 THEN '�f��' WHEN 1 THEN '����' WHEN 2 THEN '�뵲' WHEN 3 THEN T1.GatherOther END   = @TTYPE ");
            }
            if (textBox5.Text != "")
            {
                sb.Append(" and  O.BillNO=@BillNO ");
            }
            if (comboBox5.Text != "")
            {
                sb.Append(" and I.InvoiceType=@InvoiceType ");
            }
            if (textBox4.Text != "")
            {
                sb.Append(" and I.InvoiceNO=@InvoiceNO ");
            }
            if (textBox6.Text != "" && textBox9.Text != "")
            {
                sb.Append(" and  I.InvoiceDate between @InvoiceDate1 and @InvoiceDate2 ");
            }
            if (comboBox6.Text != "")
            {
                if (comboBox6.Text == "�w�}��")
                {
                    sb.Append(" and isnull(I.PrintMan,'False')='True' ");
                }
                if (comboBox6.Text == "���}��")
                {
                    sb.Append(" and isnull(I.PrintMan,'False')='False' ");
                }
            }
            if (comboBox7.Text != "")
            {
                sb.Append(" and  I.PRINTER=@PRINTER ");
            }
            if (textBox10.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  S.DEPTID BETWEEN @C1 AND @C2 ");
            }
            if (comboBox3.Text == "�q�ʳ渹")
            {
                sb.Append(" ORDER BY ID");
            }
            else if (comboBox3.Text == "�P�f�渹")
            {
                sb.Append(" ORDER BY ID2");
            }

            else if (comboBox3.Text == "�P�f���")
            {
                sb.Append(" ORDER BY �P�f���");
            }
            else if (comboBox3.Text == "�o�����X")
            {
                sb.Append(" ORDER BY �o�����X");
            }



            if (comboBox4.Text == "�ɧ�")
            {
                sb.Append(" ASC");
            }
            else if (comboBox4.Text == "����")
            {
                sb.Append(" DESC");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 3600;
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CreateDate1", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@TX", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@TTYPE", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@BillNO", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceType", comboBox5.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceNO", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceDate1", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceDate2", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@PRINTER", comboBox7.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "SALES");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData4(string BillNO)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("                              SELECT O.BillNO ID2  FROM ComProdRec O  ");
            sb.Append("                                           left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)  ");
            sb.Append("                                            left join comCustomer U On  U.ID=S.CustID  AND U.Flag =1  ");
            sb.Append("                                            Left Join comProduct B On B.ProdID =o.ProdID   ");
            sb.Append("                                                Left Join comInvoice I On  O.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1  ");
            sb.Append("                                                    Left Join StkBillSUB T0 On  O.BillNO=T0.BillNO AND  O.RowNO =T0.RowNO ");
            sb.Append("                                                       Left Join StkBillMAIN T1 On T0.BillNO=T1.BillNO  ");
            sb.Append("                                       WHERE O.Flag =500    ");
            sb.Append("                                       AND CASE T1.GatherStyle WHEN 0 THEN '�f��' WHEN 1 THEN '����' WHEN 2 THEN '�뵲' WHEN 3 THEN T1.GatherOther END <> 'FOC' ");
            sb.Append(" AND  O.BILLDATE BETWEEN @CreateDate AND @CreateDate1 AND O.BillNO=@BillNO ");

  
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CreateDate1", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetCUSTNO(string BillNO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("      SELECT CUSTNO FROM GB_CUSTINV WHERE BILLNO=@BILLNO    ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GetOrderData3TT(string BillNO)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT DISTINCT  ISNULL(G.BillNO,'')  BillNO FROM ComProdRec O    ");
            sb.Append("                                                          left join OrdBillSub G On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO    ");
            sb.Append("                                                         WHERE O.BillNO=@BillNO ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
 

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3T(string BILLNO,string DOCTYPE)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT O.ProdID �~��,B.InvoProdName �o���~�W,CAST(O.QUANTITY AS DECIMAL(10,2)) �ƶq,CASE S.TaxType WHEN 0 THEN '�t�|' WHEN 1 THEN '���|' END ���~���,O.MLPrice ���,CASE SUBSTRING(o.PRODID,1,3) WHEN 'FRE' THEN CAST(round(O.AMOUNT*1.05,0) AS INT) ELSE CAST(ISNULL(O.AMOUNT+O.TAXAMT,0) AS INT) END ���B ");
            sb.Append(" ,CASE  F.IsGift  WHEN '1' THEN 'V' END �ث~,  O.BillNO,O.RowNO,O.Flag");
            sb.Append(" FROM ComProdRec O 		 Left Join comProduct B On B.ProdID =o.ProdID ");
            sb.Append(" Left Join stkBillSub F On O.BillNO =F.BillNO and O.Flag =F.Flag AND O.RowNO =F.RowNO  LEFT JOIN comBillAccounts S ON (O.BillNO =S.FundBillNo)     ");
            if (DOCTYPE == "�P�h")
            {
                sb.Append(" WHERE O.Flag =600 ");
            }
            else if (DOCTYPE == "����")
            {
                sb.Append(" WHERE O.Flag =701 ");
            }
            else
            {
                sb.Append(" WHERE O.Flag =500 ");
            }
            sb.Append("  AND O.BILLNO=@BILLNO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData31(string BillNO)
        {


            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT   O.BillNO PO,O.ProdID INVNAME,CAST(O.QUANTITY AS DECIMAL(10,2)) QTY,CASE SUBSTRING(PRODID,1,3) WHEN 'FRE' THEN  CAST(round(O.AMOUNT*1.05,0) AS INT) ELSE CAST(ISNULL(O.AMOUNT+O.TAXAMT,0) AS INT) END AMOUNT FROM ComProdRec O  ");
            sb.Append(" WHERE  O.BillNO=@BillNO and Flag=500  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
    

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
     
        private System.Data.DataTable GetZENBEN(string BillNO)
        {


            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT BillNO  FROM ordBillMain WHERE Flag=2 AND Remark  LIKE '%�ȥ��H�e%' and BillNO =@BillNO         ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetTONBEN(string BillNO)
        {


            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT substring(Remark,CHARINDEX('�νs:#', Remark)+3,8) �νs  FROM ordBillMain WHERE Flag=2 AND Remark  LIKE '%�νs:#%' and BillNO =@BillNO         ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetTONBEN2(string BillNO)
        {


            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT REPLACE(REPLACE(REPLACE(substring(Remark,CHARINDEX('�νs', Remark)+3,8),'8.�o���a�}',''),':',''),'8.�o���a','') �νs  FROM ordBillMain WHERE Flag=2 AND Remark  LIKE '%�νs%' and BillNO =@BillNO     ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData313(string BillNO)
        {


            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT   MAX(H.CLASSID)+'   ' INVNAME,SUM(CAST(O.QUANTITY AS DECIMAL(10,1))) QTY,SUM(CAST(O.QUANTITY AS int)) QTY1,SUM(CASE SUBSTRING(O.ProdID,1,3) WHEN 'FRE' THEN  CAST(round(O.AMOUNT*1.05,0) AS INT) ELSE CAST(ISNULL(O.AMOUNT+O.TAXAMT,0) AS INT) END) AMOUNT FROM ComProdRec O  ");
            sb.Append("             Left Join comProduct H On H.ProdID=O.ProdID");
            sb.Append("           WHERE  O.BillNO=@BillNO  and Flag=500  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView1.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void POTATOAR_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
             

                // ���� PORT
                this.comport.Close();
                this.comport.Dispose();
            }
            catch { }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            EXEC();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            EXEC();
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("�п�ܦC�L���C");
                return;
            }

            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string ZfileName = lsAppDir + "\\Excel\\temp\\" + "GB.TXT";


            FileStream Zfs = new FileStream(ZfileName, FileMode.Create);
            StreamWriter Zr = new StreamWriter(Zfs);


                string F2 = dataGridView1.SelectedRows[0].Cells["ID"].Value.ToString();
                if (F2.IndexOf("/") != -1)
                {
                    F2 = F2.Substring(0, F2.IndexOf("/")) + "~" + F2.Substring(F2.LastIndexOf("/") + 1, 10);
                }
                string FF2 = dataGridView1.SelectedRows[0].Cells["ID2"].Value.ToString();
                string INV = dataGridView1.SelectedRows[0].Cells["�νs"].Value.ToString();
                string DOC = dataGridView1.SelectedRows[0].Cells["�P�f���"].Value.ToString();
                string ���|���B = dataGridView1.SelectedRows[0].Cells["���|���B"].Value.ToString();

                string �Ȥ�W�� = dataGridView1.SelectedRows[0].Cells["�Ȥ�W��"].Value.ToString();
                string �ҵ|���O = dataGridView1.SelectedRows[0].Cells["�ҵ|���O"].Value.ToString();
                int MAN = �Ȥ�W��.IndexOf("�֪�");
                int MAN2 = �Ȥ�W��.IndexOf("�E���t�a");
                int MAN3 = �Ȥ�W��.IndexOf("�ڷ�");
                System.Data.DataTable T1 = GetOrderData31(FF2);
      
                string DOCDATE = DOC.Substring(0, 4) + "/" + DOC.Substring(4, 2) + "/" + DOC.Substring(6, 2);

                Zr.WriteLine("���׹�~�ѥ��������q" + System.Environment.NewLine);
                 Zr.WriteLine("��~�H�νs: 22468373" + System.Environment.NewLine);
                 Zr.WriteLine("�x�_������Ϸs��G��" + System.Environment.NewLine);
                 Zr.WriteLine("257��5�Ӥ�3 TEL:87922800" + System.Environment.NewLine);
                 Zr.WriteLine("POS# ARMAS-001" + System.Environment.NewLine);
                 Zr.WriteLine(DOCDATE + System.Environment.NewLine);
                if (!String.IsNullOrEmpty(INV))
                {
                    Zr.WriteLine("�Τ@�s��: " + INV + System.Environment.NewLine);
                 
                }
                 Zr.WriteLine("------------------------" + System.Environment.NewLine);


                if (T1.Rows.Count > 0)
                {
                    if (T1.Rows.Count > 9 || ���|���B.Length >4)
                    {
                        System.Data.DataTable TT1 = GetOrderData313(FF2);
                        
                        string INVNAME = TT1.Rows[0]["INVNAME"].ToString();
                       
                        string QTY = "";
                        if (MAN != -1 || MAN2 != -1 || MAN3 != -1)
                        {
                    
                            QTY = TT1.Rows[0]["QTY1"].ToString();
                        }
                        else
                        {
                            QTY = TT1.Rows[0]["QTY"].ToString();
                        }
                        int QTYT = QTY.Length;
                        if (QTYT == 1)
                        {
                            QTY = "    " + QTY;
                        }
                        else if (QTYT == 2)
                        {
                            QTY = "   " + QTY;
                        }
                        else if (QTYT == 3)
                        {
                            QTY = "  " + QTY;
                        }
                        else if (QTYT == 4)
                        {
                            QTY = " " + QTY;
                        }
              
                        int TAMOUNT = Convert.ToInt32(TT1.Rows[0]["AMOUNT"].ToString());
                        string AMOUNT = TAMOUNT.ToString("#,##0");
                        int AMOUNTT = AMOUNT.Length;
                        if (AMOUNTT == 1)
                        {
                            AMOUNT = "       " + AMOUNT;
                        }
                        if (AMOUNTT == 2)
                        {
                            AMOUNT = "      " + AMOUNT;
                        }
                        if (AMOUNTT == 3)
                        {
                            AMOUNT = "     " + AMOUNT;
                        }
                        if (AMOUNTT == 4)
                        {
                            AMOUNT = "    " + AMOUNT;
                        }
                        if (AMOUNTT == 5)
                        {
                            AMOUNT = "   " + AMOUNT;
                        }
                        if (AMOUNTT == 6)
                        {
                            AMOUNT = "  " + AMOUNT;
                        }
                        if (AMOUNTT == 7)
                        {
                            AMOUNT = " " + AMOUNT;
                        }


                        Zr.WriteLine(INVNAME + QTY + AMOUNT + �ҵ|���O + System.Environment.NewLine);
                    }
                    else
                    {
                        for (int i = 0; i <= T1.Rows.Count - 1; i++)
                        {
                            string INVNAME = T1.Rows[i]["INVNAME"].ToString();
                        
                            string QTY = T1.Rows[i]["QTY"].ToString();
                            int QTYT = QTY.Length;
                            if (QTYT == 1)
                            {
                                QTY = "     " + QTY;
                            }
                            else if (QTYT == 2)
                            {
                                QTY = "    " + QTY;
                            }
                            else if (QTYT == 3)
                            {
                                QTY = "   " + QTY;
                            }
                            else if (QTYT == 4)
                            {
                                QTY = "  " + QTY;
                            }
                            else if (QTYT == 5)
                            {
                                QTY = " " + QTY;
                            }


                            int TAMOUNT = Convert.ToInt32(T1.Rows[i]["AMOUNT"].ToString());
                            string AMOUNT = TAMOUNT.ToString("#,##0");
                            int AMOUNTT = AMOUNT.Length;
                            if (AMOUNTT == 5)
                            {
                                AMOUNT = "  " + AMOUNT;
                            }
                            if (AMOUNTT == 6)
                            {
                                AMOUNT = " " + AMOUNT;
                            }
                            if (AMOUNTT == 3)
                            {
                                AMOUNT = "    " + AMOUNT;
                            }
                            if (AMOUNTT == 2)
                            {
                                AMOUNT = "     " + AMOUNT;
                            }
                            if (AMOUNTT == 1)
                            {
                                AMOUNT = "      " + AMOUNT;
                            }


                            Zr.WriteLine(INVNAME + QTY + AMOUNT + �ҵ|���O + System.Environment.NewLine);
                            //  }

                        }
                    }


                    int T���B = Convert.ToInt32(���|���B);
                    string ���B = T���B.ToString("#,##0");
                    string ���B2 = T���B.ToString("#,##0");
                    int ���BT = ���B.Length;
                    string �K�|���B = "";
                    if (���BT == 5)
                    {
                        ���B = "            " + ���B;
                    }
                    if (���BT == 6)
                    {
                        ���B = "           " + ���B;
                    }
                    if (���BT == 7)
                    {
                        ���B = "          " + ���B;
                    }
                    if (���BT == 3)
                    {
                        ���B = "              " + ���B;
                    }

                    if (���BT == 5)
                    {
                        �K�|���B = "          " + ���B2;
                    }
                    if (���BT == 6)
                    {
                        �K�|���B = "         " + ���B2;
                    }
                    if (���BT == 7)
                    {
                        �K�|���B = "        " + ���B2;
                    }
                    if (���BT == 3)
                    {
                        �K�|���B = "            " + ���B2;
                    }


                    int T�`�p = Convert.ToInt32(���|���B);
                    string �`�p = T�`�p.ToString("#,##0");
                    int �`�pT = �`�p.Length;
                    if (�`�pT == 5)
                    {
                        �`�p = "              " + �`�p;
                    }
                    if (�`�pT == 6)
                    {
                        �`�p = "             " + �`�p;
                    }
                    if (�`�pT == 7)
                    {
                        �`�p = "            " + �`�p;
                    }
                    if (�`�pT == 3)
                    {
                        �`�p = "                " + �`�p;
                    }
                     Zr.WriteLine("------------------------" + System.Environment.NewLine);
                     Zr.WriteLine("�p�p:" + ���B + �ҵ|���O + System.Environment.NewLine);

                     Zr.WriteLine("========================" + System.Environment.NewLine);
                     Zr.WriteLine("�`�p:" + �`�p + System.Environment.NewLine);

      
                     Zr.WriteLine("" + System.Environment.NewLine);
                     Zr.WriteLine("PO# " + F2 + System.Environment.NewLine);
                     Zr.WriteLine("SO# " + FF2 + System.Environment.NewLine);
                     Zfs.Flush();
                     Zr.Close();
                     System.Diagnostics.Process.Start(ZfileName);
                
            }
        }
        public void INSERTINVOICE(string FLAG, string InvoBillNo, string InvoiceNO, string InvoiceDate, string SrcSysNO, string SrcBillNO, string SrcBillFlag, string CustomerID, string ApplyMonth, string InvoiceType, string OffsetType, int TaxType, string OtherVoucher, decimal Amount, decimal TaxAmt, string CompanyName, string TaxRegNO, string ZipCode, int UseOrder, int IncludeTax, string Remark, string InvoMonth, decimal Total, string InvoRealMonth, string Printed, string PrintMan, string PrintDate, string Printer, string SpecialTaxType, string Address, string InvoAddr, string InvoPool, string InvoPoolEndNo, string InvoAlcoholandSmoke, string IsCancel, string CancelType, string IsShowTax, string MakerID, string Maker, string ReportCompID)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO COMINVOICE (FLAG,InvoBillNo,InvoiceNO,InvoiceDate,SrcSysNO,SrcBillNO,SrcBillFlag,CustomerID,ApplyMonth,InvoiceType,OffsetType,TaxType,OtherVoucher,Amount,TaxAmt,CompanyName,TaxRegNO,ZipCode,UseOrder,IncludeTax,Remark,InvoMonth,Total,InvoRealMonth,Printed,PrintMan,PrintDate,Printer,SpecialTaxType,Address,InvoAddr,InvoPool,InvoPoolEndNo,InvoAlcoholandSmoke,IsCancel,CancelType,IsShowTax,MakerID,Maker,ReportCompID,MergeOutState) VALUES(@FLAG,@InvoBillNo,@InvoiceNO,@InvoiceDate,@SrcSysNO,@SrcBillNO,@SrcBillFlag,@CustomerID,@ApplyMonth,@InvoiceType,@OffsetType,@TaxType,@OtherVoucher,@Amount,@TaxAmt,@CompanyName,@TaxRegNO,@ZipCode,@UseOrder,@IncludeTax,@Remark,@InvoMonth,@Total,@InvoRealMonth,@Printed,@PrintMan,@PrintDate,@Printer,@SpecialTaxType,@Address,@InvoAddr,@InvoPool,@InvoPoolEndNo,@InvoAlcoholandSmoke,@IsCancel,@CancelType,@IsShowTax,@MakerID,@Maker,@ReportCompID,@MergeOutState)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@FLAG", FLAG));
            command.Parameters.Add(new SqlParameter("@InvoBillNo", InvoBillNo));
            command.Parameters.Add(new SqlParameter("@InvoiceNO", InvoiceNO));
            command.Parameters.Add(new SqlParameter("@InvoiceDate", InvoiceDate));
            command.Parameters.Add(new SqlParameter("@SrcSysNO", SrcSysNO));
            command.Parameters.Add(new SqlParameter("@SrcBillNO", SrcBillNO));
            command.Parameters.Add(new SqlParameter("@SrcBillFlag", SrcBillFlag));
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@ApplyMonth", ApplyMonth));
            command.Parameters.Add(new SqlParameter("@InvoiceType", InvoiceType));
            command.Parameters.Add(new SqlParameter("@OffsetType", OffsetType));
            command.Parameters.Add(new SqlParameter("@TaxType", TaxType));
            command.Parameters.Add(new SqlParameter("@OtherVoucher", OtherVoucher));
            command.Parameters.Add(new SqlParameter("@Amount", Amount));
            command.Parameters.Add(new SqlParameter("@TaxAmt", TaxAmt));
            command.Parameters.Add(new SqlParameter("@CompanyName", CompanyName));
            command.Parameters.Add(new SqlParameter("@TaxRegNO", TaxRegNO));
            command.Parameters.Add(new SqlParameter("@ZipCode", ZipCode));
            command.Parameters.Add(new SqlParameter("@UseOrder", UseOrder));
            command.Parameters.Add(new SqlParameter("@IncludeTax", IncludeTax));
            command.Parameters.Add(new SqlParameter("@Remark", Remark));
            command.Parameters.Add(new SqlParameter("@InvoMonth", InvoMonth));
            command.Parameters.Add(new SqlParameter("@Total", Total));
            command.Parameters.Add(new SqlParameter("@InvoRealMonth", InvoRealMonth));
            command.Parameters.Add(new SqlParameter("@Printed", Printed));
            command.Parameters.Add(new SqlParameter("@PrintMan", PrintMan));
            command.Parameters.Add(new SqlParameter("@PrintDate", PrintDate));
            command.Parameters.Add(new SqlParameter("@Printer", Printer));
            command.Parameters.Add(new SqlParameter("@SpecialTaxType", SpecialTaxType));
            command.Parameters.Add(new SqlParameter("@Address", Address));
            command.Parameters.Add(new SqlParameter("@InvoAddr", InvoAddr));
            command.Parameters.Add(new SqlParameter("@InvoPool", InvoPool));
            command.Parameters.Add(new SqlParameter("@InvoPoolEndNo", InvoPoolEndNo));
            command.Parameters.Add(new SqlParameter("@InvoAlcoholandSmoke", InvoAlcoholandSmoke));
            command.Parameters.Add(new SqlParameter("@IsCancel", IsCancel));
            command.Parameters.Add(new SqlParameter("@CancelType", CancelType));
            command.Parameters.Add(new SqlParameter("@IsShowTax", IsShowTax));
            command.Parameters.Add(new SqlParameter("@MakerID", MakerID));
            command.Parameters.Add(new SqlParameter("@Maker", Maker));
            command.Parameters.Add(new SqlParameter("@ReportCompID", ReportCompID));
            command.Parameters.Add(new SqlParameter("@MergeOutState", "0"));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }

        private void UpdateID2(string ID, string PrintMan, string InvoiceNO)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            if (String.IsNullOrEmpty(ID))
            {
                sb.Append(" UPDATE   comInvoice SET PrintMan=@PrintMan WHERE InvoiceNO=@InvoiceNO");
            }
            else
            {
                sb.Append(" UPDATE   comInvoice SET PrintMan=@PrintMan WHERE SrcBillNO=@ID");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@PrintMan", PrintMan));
            command.Parameters.Add(new SqlParameter("@InvoiceNO", InvoiceNO));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }

        private void UpdateCUSTINV(string TaxRegNO, string SrcBillNO)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE   comInvoice SET TaxRegNO=@TaxRegNO where SrcBillNO=@SrcBillNO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@TaxRegNO", TaxRegNO));
            command.Parameters.Add(new SqlParameter("@SrcBillNO", SrcBillNO));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
  
        private void UpdateID3(string ID, string PRINTER)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
       
                sb.Append(" UPDATE   comInvoice SET PRINTER=@PRINTER WHERE SrcBillNO=@ID");
            

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@PRINTER", PRINTER));
     
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 0)
            {
                string �P�f�渹 = dataGridView1.SelectedRows[0].Cells["ID2"].Value.ToString();
                string ID = dataGridView1.SelectedRows[0].Cells["ID"].Value.ToString();
                string DOCTYPE = "";
                if (comboBox8.Text != "")
                {
                    DOCTYPE = comboBox8.Text;
                }
                System.Data.DataTable dtT = GetOrderData3T(�P�f�渹, DOCTYPE);
                dataGridView3.DataSource = dtT;


                DataRow rowT;
                //�[�J�@���X�p
                Decimal[] TotalT = new Decimal[dtT.Columns.Count - 1];

                for (int i = 0; i <= dtT.Rows.Count - 1; i++)
                {

                    for (int j = 2; j <= 5; j++)
                    {
                        if (j != 3)
                        {
                            try
                            {
                                TotalT[j - 1] += Convert.ToDecimal(dtT.Rows[i][j]);
                            }
                            catch
                            {
                                TotalT[j - 1] += 0;
                            }
                        }

                    }
                }



                rowT = dtT.NewRow();

                rowT[1] = "�X�p";
                for (int j = 2; j <= 5; j++)
                {
                    if (j != 3 && j != 4)
                    {
                        rowT[j] = TotalT[j - 1];
                    }
                }
                dtT.Rows.Add(rowT);

                for (int i = 2; i <= 5; i++)
                {
                    DataGridViewColumn col = dataGridView3.Columns[i];


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    if (i == 2)
                    {
                        col.DefaultCellStyle.Format = "#,##0.00";
                    }

          
                    if (i == 5)
                    {
                        col.DefaultCellStyle.Format = "#,##0";
                    }
                }
                if (ID == "���J��")
                {
                    System.Data.DataTable K1 = GetREMARK2(�P�f�渹);
                    if (K1.Rows.Count > 0)
                    {
                        textBox3.Text = K1.Rows[0][0].ToString();

                    }
                }
                else
                {
                    System.Data.DataTable K1 = GetREMARK(�P�f�渹, comboBox8.Text);
                    if (K1.Rows.Count > 0)
                    {
                        textBox3.Text = K1.Rows[0][0].ToString();

                    }
                }
            }
        }

        private void dataGridView3_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView3.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView3.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
            //�ث~
      //      sb.Append(" SELECT  O.ProdID �~��,B.InvoProdName �o���~�W,CAST(O.QUANTITY AS DECIMAL(10,2)) �ƶq,CASE S.TaxType WHEN 0 THEN '�t�|' WHEN 1 THEN '���|' END ���~���,O.MLPrice ���,CASE SUBSTRING(o.PRODID,1,3) WHEN 'FRE' THEN CAST(round(O.AMOUNT*1.05,0) AS INT) ELSE CAST(ISNULL(O.AMOUNT+O.TAXAMT,0) AS INT) END ���B ");
            if (dgr.Cells["���B"].Value.ToString() == "0" && dgr.Cells["�ث~"].Value.ToString() != "V")
            {

                dgr.DefaultCellStyle.ForeColor = Color.Red;
            }

        }


        private void button2_Click(object sender, EventArgs e)
        {
            for (int j = 0; j <= dataGridView1.Rows.Count - 2; j++)
            {
                string FF3 = dataGridView1.Rows[j].Cells["PRINT"].Value.ToString();
                string FF2 = dataGridView1.Rows[j].Cells["ID2"].Value.ToString();
                string �o�����X = dataGridView1.Rows[j].Cells["�o�����X"].Value.ToString();
                
                if (FF3 == "True")
                {
                    UpdateID2(FF2, "True", �o�����X);
                }
                else
                {
                    UpdateID2(FF2, "False", �o�����X);
                }
            }
            MessageBox.Show("��Ƥw��s");

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
                      DialogResult result;
            result = MessageBox.Show("�N��s��ƽнT�w�O�_�n��s", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                if (checkBox2.Checked)
                {

                    for (int j = 0; j <= dataGridView1.Rows.Count - 1; j++)
                    {
                        string FF3 = dataGridView1.Rows[j].Cells["PRINT"].Value.ToString();
                        string FF2 = dataGridView1.Rows[j].Cells["ID2"].Value.ToString();
                        string �o�����X = dataGridView1.Rows[j].Cells["�o�����X"].Value.ToString();
                        UpdateID2(FF2, "True", �o�����X);
                    }
                }
                else
                {
                    for (int j = 0; j <= dataGridView1.Rows.Count - 1; j++)
                    {
                        string FF3 = dataGridView1.Rows[j].Cells["PRINT"].Value.ToString();
                        string FF2 = dataGridView1.Rows[j].Cells["ID2"].Value.ToString();
                        string �o�����X = dataGridView1.Rows[j].Cells["�o�����X"].Value.ToString();
                        UpdateID2(FF2, "False", �o�����X);
                    }
                }

                MessageBox.Show("��Ƥw��s");

            }
            checkBox2.Checked = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
    
            MAKE("0","Y");
 
        }
        //string host = "61.57.227.80";
        //string username = "22468373p";
        //string password = "b152224$P";
        private bool isValidConnection(string url, string user, string password)
        {
            FtpWebRequest request;
            try
            {
                request = (FtpWebRequest)FtpWebRequest.Create(@"FTP://" + url);
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.KeepAlive = true;
                request.Credentials = new NetworkCredential(user, password);
                var response = request.GetResponse();
                response.Close();
            }
            catch (WebException ex)
            {
                return false;
            }
            finally
            {
                request = null;
            }

            return true;
        }
        private void MAKE(string STATUS,string FLAG)
        {
            string OrderNo = "";
            string OrderNo2 = "";
            if (FLAG == "Y")
            {
                if (MessageBox.Show("�T�w����ܡH", "�H������", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    return;
                }

            }
            DataGridViewRow row;
            for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = dataGridView1.SelectedRows[i];

                OrderNo = Convert.ToString(row.Cells["ID2"].Value);
                int �K�|���B = Convert.ToInt32(row.Cells["�K�|���B"].Value);
                int ���|���B = Convert.ToInt32(row.Cells["���|���B"].Value);
                int �|�B = Convert.ToInt32(row.Cells["�|�B"].Value);
                 //�|�B
                 if (�|�B != 0)
                 {
                     int G1 = Convert.ToInt32(���|���B / 1.05);
                     if (G1 != �K�|���B)
                     {
                       
                             MessageBox.Show("�P�f�渹: " + OrderNo + " �|�B���ŦX���ˬd");
                             return;
                         
                     }
                 }
            }




            string InvoiceFileName = string.Format("{0}-O-{1}-{2}.txt", FirmNo, DateTime.Now.ToString("yyyyMMdd"), DateTime.Now.ToString("HHmmss"));

            FileStart(InvoiceFileName);





            for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = dataGridView1.SelectedRows[i];

                OrderNo = Convert.ToString(row.Cells["ID2"].Value);
                OrderNo2 = Convert.ToString(row.Cells["ID"].Value);
                string �νs = Convert.ToString(row.Cells["�νs"].Value);

                string �����ȥ� = Convert.ToString(row.Cells["�����ȥ�"].Value);
                MakeData(OrderNo, STATUS, �νs, �����ȥ�, OrderNo2);
                string A1 = "";
           
                if (STATUS == "0")
                {
                    A1 = "�s�W";
                }
                if (STATUS == "1")
                {
                    A1 = "�׳�";
                }
                if (STATUS == "3")
                {
                    A1 = "�P�h";
                }
                if (STATUS == "4")
                {
                    A1 = "�P��";
                }
                if (FLAG == "Y")
                {
                    UpdateID3(OrderNo, A1);
                }

            }


            FileClose(dataGridView1.SelectedRows.Count.ToString());

            UploadToFtp(InvoiceFileName,FLAG);
        }
        private void UploadToFtp(string InvoiceFileName,string FLAG)
        {
          
            try
            {
                string Msg = "";

                //  string OrderFileName = "TestFile1.txt";
                string OrderFileName = InvoiceFileName;


               FTPclient ftp = new FTPclient(host, username, password);

               //GU FF = new GU(host2, username, password);
                //Receive //Send
                //List<string> l = ftp.ListDirectory("/");


                string FileName = GetExePath() + "\\EXCEL\\temp\\" + OrderFileName;


                //ftp.Upload(FileName, UpLoadDataPath + DateTime.Now.ToString("HHmmss") + OrderFileName);
                if (FLAG == "Y")
                {
                    ftp.Upload(FileName, UpLoadDataPath + OrderFileName);
                    
                  
                    Msg = "�W�Ǧ��\";

                    MessageBox.Show(Msg);

                    DELETEFILE2();

                }
                else
                {
                    System.Diagnostics.Process.Start("notepad.exe", FileName);
                }

            }
            catch
            {

            }
        }

      
        private void FileStart(string InvoiceFileName)
        {

            sw = new StreamWriter(GetExePath() + "\\Excel\\temp\\" + InvoiceFileName, false, Encoding.UTF8);//creating html file

        }
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }
        private void FileClose(string RecordCount)
        {

            string DEnd = RecordCount + CrLf;

            sw.Write(DEnd);
            sw.Close();
        }

        private void MakeData(string OrderNumber, string FF, string �νs, string �����ȥ�, string OrderNo2)
        {
            //���ն��q�A�o���}�߹L�{�|�o�eEmail�����q���H

            string MailTest = "acmegb-fin@acmegb.com";
       //     string MailTest = "lleytonchen@acmepoint.com";
            //���ո��
            string sqlMaster = "select * from comBillAccounts where Flag=500 and FundBillNO ='{0}'";
     
            string sqlDetail = "select * from comProdRec where Flag=500 and BillNO ='{0}'";

            string sqlF = " SELECT Remark  FROM comInvoice WHERE SrcBillNO ='{0}' AND ISNULL(Remark,'') <> ''";
            if (FF == "3")
            {
                sqlMaster = "select * from comBillAccounts where Flag IN  (600)  and FundBillNO ='{0}'";
                sqlDetail = "select * from comProdRec where Flag IN  (600) and BillNO ='{0}'";
            }
            if (FF == "4")
            {
                sqlMaster = "SELECT MAX(T0.DISTNO)  FundBillNO,MAX(DispBillDate) BillDate,CASE WHEN SUM(T1.TaxAmt ) =0 THEN 3 ELSE 1 END TaxType,SUM(T1.Dist) Total ,SUM(T1.TaxAmt ) Tax,'' CustID,'' ZipCode,'' Remark,MAX(T1.CustBillNo ) CustBillNo FROM StkDistmain T0 LEFT JOIN stkDistSub T1 ON (T0.DISTNO =T1.DistNO) WHERE T0.DistNO ='{0}' ";
                sqlDetail = "select BillNO,ProdID,ProdName,Price,Quantity,mldist Amount,CASE WHEN TaxAmt=0 THEN 0 ELSE  round(mldist*0.05,0) END TaxAmt,* from comProdRec where Flag IN  (701) and BillNO ='{0}'";
            }
            if (OrderNo2 == "���J��")
            {
                sqlMaster = "select * from comBillAccounts where Flag IN  (595)  and FundBillNO ='{0}'";
                sqlDetail = "Select CostBillNo BillNO,ItemNo ProdID,FareClassName ProdName,[Money] Price,1 Quantity,[Money]  Amount,[Money]+MLTaxAmt  TaxAmt    From comCostSub A Left Join comFareMeans B On B.Flag=A.Flag-79 And B.FareClassID=A.ItemNo Where A.Flag=80 And A.CostBillNo='{0}'";
            }
            DataTable dtMaster = GetData(string.Format(sqlMaster, OrderNumber));

            DataTable dtDetail = GetData(string.Format(sqlDetail, OrderNumber));
            DataTable dtF = GetData(string.Format(sqlF, OrderNumber));
            //gvMaster.DataSource = dtMaster;

            //gvDetail.DataSource = dtDetail;


            //�@���q��@���ɮ�
            //�q���� �ɦW�G�νs-O-yyyymmdd-hhmmss.txt
            //���j�Ÿ��� ���j�Ÿ��� |�AUNICODE-UTF-8�s�X
            //���� CrLf = "\r\n"

            //�W�Ǯ榡
            //�ǿ��Ʈ榡�G
            //M
            //D
            //D
            //D

            //�D�ɳW��
            //F01 �D�ɥN��* M
            string F01 = "M";
            //F02 �q��s��* C40
            string F02 = "";
            //F03 �q�檬�A* 0 :�s�W 1: �׳� (�������h�f �������h�f ) 2: �R
            string F03 = "0"; //�s�W
            //F04 �q����* C10 : 2010/11/15
            string F04 = "2016/01/01";
            //F05 �w�p�X�f���* C10
            //�ѩ�o���Ǯɸ��A�]��������
            string F05 = "2016/01/01";
            //F06 �|�v�O* C1 1:���| 2:�s�|�v 3:�K�|
            string F06 = "";
            //F07 �q����B N14 (���|)
            string F07 = "";
            //F08 �q��|�B N14
            string F08 = "";
            //F09 �q����B(�t�|)* N14
            string F09 = "";
            //F10 ���Τ@�s��-C8
            string F10 = "";
            //F11 ���t�s*-C20
            string F11 = "";
            //F12 �R��Τ@�s�� -C8
            string F12 = "";
            //F13 �R���H���q�W�� C160
            string F13 = "";
            //F14 �|���s��* C40 
            string F14 = "";
            //F15 �|���m�W* C80
            string F15 = "";
            //F16 �|���l���ϸ� C5 
            string F16 = "";
            //F17 �|���a�}* C240 
            string F17 = "";
            //F18 �|���q�� C20
            string F18 = "";
            //F19 �|����ʹq�� C20
            string F19 = "";
            //F20 �|���q�l�l��* C100
            string F20 = "";
            //F21 ���Q�I�Ƨ馩���B N14
            string F21 = "";
            //F22 �����ȥ��o��* C1 Y: �ȥ� N: �D�ȥ�
            string F22 = "N";
      
            //F23 �o�����ص��O C20
            string F23 = "";
            //F24 �q����O C20
            string F24 = "";
            //F25 �I�ڤ覡 C100
            string F25 = "";
            //F26 �������X 1 C20 ( �X�f�渹 �X�f�渹 )
            string F26 = "";
            //F27 �������X 2 C20
            string F27 = "";
            //F28 �������X 3 C20
            string F28 = "";
            //F29 �D�ɳƵ� C100
            string F29 = "";
            //F30 �ӫ~�W�� C100 (�ȥ��o���W�L 10 �~�����ȫh �Ѧ��W�٥N��)
            string F30 = "";
            //F31 �������O���X C6
            string F31 = "";
            //F32 ������X id1( ���X ) C64
            string F32 = "";
            //F33 �������X id2( ���X ) C64
            string F33 = "";


            //BindData
            DataRow dr = dtMaster.Rows[0];

            DataRow dr2 = dtDetail.Rows[0];


            //F01 �D�ɥN��* M
            F01 = "M";
            //F02 �q��s��* C40
            BindData(dr, "FundBillNO", ref F02);
            if (FF == "3")
            {
                BindData(dr2, "FromNO", ref F02);
                BindData(dr, "FundBillNO", ref F27);
            }
            if (FF == "4")
            {
                BindData(dr, "CustBillNo", ref F02);
                BindData(dr, "CustBillNo", ref F27);
            }
            //if (FF == "2")
            //{
            //    if (dtF.Rows.Count > 0)
            //    {
            //        F02 = dtF.Rows[0][0].ToString();

            //        //DataTable S = GetData(string.Format(sqlF2, F02));
            //        //if (S.Rows.Count == 0)
            //        //{
            //        //    MessageBox.Show("�W�ǥ��ѡA�S����" + F02 + " �P�f�u��");
            //        //    return;
            //        //}
            //    }
            //    else
            //    {
            //        MessageBox.Show("�W�ǥ��ѡA�S�����o��");
            //        return;
            //    }
                
            //}
            
            //F03 �q�檬�A* 0 :�s�W 1: �׳� (�������h�f �������h�f ) 2: �R
            F03 = FF; //�s�W
            if (FF == "4")
            {
                F03 = "3";
            }
            ////F04 �q����* C10 : 2010/11/15
            BindData(dr, "BillDate", ref F04);
            F04 = ConvertDate(F04);
            ////F05 �w�p�X�f���* C10
            ////�ѩ�o���Ǯɸ��A�]��������
            //BindData(dr, F05 = "2016/01/01";
            //�b������-�����n�����D��
            try
            {
                //��ڨ�f��
                //BindData(dr,"UDef1" ,ref F05);
                BindData(dr, "BillDate", ref F05);

               

                F05 = ConvertDate(F05);
            }
            catch
            {
            }

            ////F06 �|�v�O* C1 1:���| 2:�s�|�v 3:�K�|
            //TaxType=0  ���|
            //TaxType=1  �K�|
            string TaxType = Convert.ToString(dr["TaxType"]);
            F06 = "3";
            if (TaxType == "0")
            {
                F06 = "1";
            }
            //BindData(dr, F06="";
            ////F07 �q����B N14 (���|)
            BindData(dr, "Total", ref F07, "");

            ////F08 �q��|�B N14
            BindData(dr, "Tax", ref F08, "");

            ////F09 �q����B(�t�|)* N14
            //BindData(dr, F09 = "";
            F09 = Convert.ToString(Convert.ToInt32(F07) + Convert.ToInt32(F08));
            ////F10 ���Τ@�s��-C8
            //BindData(dr, F10 = "";
            F10 = FirmNo;

            ////F11 ���t�s*-C20
            //BindData(dr, F11 = "";
            ////F12 �R��Τ@�s�� -C8
            //BindData(dr, F12 = "";
            ////F13 �R���H���q�W�� C160
            //BindData(dr, F13 = "";

            F12 = �νs.Trim();
            ////F14 �|���s��* C40 
            BindData(dr, "CustID", ref F14);
            ////F15 �|���m�W* C80

            string email = MailTest;
            string MobileTel = "";
            string CLASSID = "";
            DataTable dtC = GetData(string.Format("select Fullname,email,MobileTel,CLASSID from comCustomer where flag=1 and  id='{0}'", F14));
            F15 = F14;
            if (dtC.Rows.Count > 0)
            {
                F15 = Convert.ToString(dtC.Rows[0]["Fullname"]);
                CLASSID = Convert.ToString(dtC.Rows[0]["CLASSID"]);
                //���ծ�,���� ���ժ�
                try
                {
                    MobileTel = Convert.ToString(dtC.Rows[0]["MobileTel"]);
                }
                catch
                {
                }
            }
            if (OrderNo2 == "���J��")
            {
                DataTable dtI = GetData(string.Format(" SELECT CompanyName COMPANY FROM comInvoice WHERE SrcBillNO ='{0}'", OrderNumber));
                if (dtI.Rows.Count > 0)
                {
                    F15 = Convert.ToString(dtI.Rows[0]["COMPANY"]);
                }
            }
            ////F16 �|���l���ϸ� C5 

            BindData(dr, "ZipCode", ref F16);
            ////F17 �|���a�}* C240 
            BindData(dr, "Remark", ref F17);
            string RR = F17;
            if (�����ȥ� == "Y")
            {
                F22 = "Y";
            }
                int G1 = F17.IndexOf("8.�o���a�}:");
                if (G1 != -1)
                {
                    string GS = F17;
                    int G2 = F17.IndexOf("9.�q�ʤHEmail:");

                    if (G2 != -1)
                    {
                        F17 = F17.Substring(G1, G2 - G1 - 1).Replace("8.�o���a�}:", "").Replace("\r", "").Trim();

                       

                    }
                    else
                    {
                        F17 = F17.Substring(G1, F17.Length - G1).Replace("8.�o���a�}:", "").Replace("\r", "").Replace("10.�O�_��DM:�O", "").Replace("\n", "").Trim();
                    }

             
                }
                else
                {
                    F17 = "";
                }

                int GG2 = RR.IndexOf("9.�q�ʤHEmail:");
                if (GG2 != -1)
                {
                    string GS = RR;
                    string E1 = GS.Substring(GG2, GS.Length - GG2).Replace("9.�q�ʤHEmail:", "");
                    int G4 = E1.IndexOf("\r");
                    if (G4 != -1)
                    {
                        email = E1.Substring(0, G4).Trim();
                    }
                    else
                    {
                        email = E1.Trim();

                    }
                }
            ////F18 �|���q�� C20
            //BindData(dr, F18 = "";
            ////F19 �|����ʹq�� C20
            F19 = MobileTel;

            ////F20 �|���q�l�l��* C100
            //BindData(dr, F20 = "";
   

            ////F21 ���Q�I�Ƨ馩���B N14
            //BindData(dr, F21 = "";

            ////F22 �����ȥ��o��* C1 Y: �ȥ� N: �D�ȥ�
            //BindData(dr, F22="N";
            ////F23 �o�����ص��O C20
            //BindData(dr, F23 = "";
            ////F24 �q����O C20
            //BindData(dr, F24 = "";
            ////F25 �I�ڤ覡 C100
            //BindData(dr, F25 = "";
            ////F26 �������X 1 C20 ( �X�f�渹 �X�f�渹 )
            //BindData(dr, F26 = "";
            ////F27 �������X 2 C20
            //BindData(dr, F27 = "";
            ////F28 �������X 3 C20
            //BindData(dr, F28 = "";
            ////F29 �D�ɳƵ� C100




            try
            {
                string remark = Convert.ToString(dr["remark"]);

                string[] sArray = remark.Split('\r');
                    int F2 = 0;
                    foreach (string F in sArray)
                    {
                        F2++;
                    }
                    //if (F2 > 1)
                    //{
                    //    string tmpOrder = sArray[2];

                    //    string[] sArray1 = tmpOrder.Split(':');
                    //    string H1 = sArray1[1];
                    //    string H2 = "";
                    //    if (!String.IsNullOrEmpty(H1))
                    //    {
                    //        System.Data.DataTable T1 = GetCARD(H1);
                    //        if (T1.Rows.Count > 0)
                    //        {
                    //            H2 = " �d�����|�X:" + T1.Rows[0][0].ToString();
                    //        }

                    //        F29 = "���ʳ渹:" + H1 + H2;
                    //    }



                    //}


                    if (F2 > 2)
                    {
                        string tmpOrder = sArray[2];

                        string[] sArray1 = tmpOrder.Split(':');
                        string H1 = sArray1[1];
                        string H2 = "";
                        string HH = "";
                        if (!String.IsNullOrEmpty(H1))
                        {
                            System.Data.DataTable T1 = GetCARD(H1);
                            if (T1.Rows.Count > 0)
                            {
                                HH = T1.Rows[0][0].ToString().Trim();
                                H2 = " �d�����|�X:" + HH;

                            }
                            F29 = "���ʳ渹:" + H1.Trim() + H2;
                        }


                        if (string.IsNullOrEmpty(HH))
                        {
                            if (F2 > 9)
                            {
                                string tmpOrder2 = sArray[9];
                                int INT1 = tmpOrder2.IndexOf("�d�����|�X");
                                if (INT1 != -1)
                                {
                                    string[] sArray12 = tmpOrder2.Split(':');
                                    string H3 = sArray12[1];
                                    H2 = " �d�����|�X:" +  H3.ToString().Trim();
                                }
                                F29 = "���ʳ渹:" + H1.Trim() + H2;
                            }
                        }
                    }

                
            }
            catch
            {

            }

            if (String.IsNullOrEmpty(email))
            {

                System.Data.DataTable GG1 = GETCLASSID(CLASSID);
                if (GG1.Rows.Count > 0)
                {
                    email = "acmegb-fin@acmegb.com";
                }
            }

            F20 = email;
            string ProdName = "";
            string QTY = "";
            string PRICE = "";
            string MARK = "";
            string EMAIL = "";
            if (OrderNo2 == "���J��")
            {

                System.Data.DataTable GG1 = GetREMARK2(OrderNumber);
                if (GG1.Rows.Count > 0)
                {
                    string REMARK = GG1.Rows[0][0].ToString();
                    int AG1 = REMARK.IndexOf("1.�~�W:");
                    if (AG1 != -1)
                    {
                        string GS = REMARK;
                        int G2 = REMARK.IndexOf("2.�ƶq:");
                        int G3 = REMARK.IndexOf("3.���:");
                        int G4 = REMARK.IndexOf("4.�Ƶ�:");
                        int G5 = REMARK.ToUpper().IndexOf("5.E-MAIL:");
                        //5.E-mail:
                        if (G2 != -1)
                        {
                            ProdName = REMARK.Substring(AG1, G2 - AG1 - 1).Replace("1.�~�W:", "").Replace("\r", "");
                            QTY = REMARK.Substring(G2, G3 - G2 - 1).Replace("2.�ƶq:", "").Replace("\r", "");
                            PRICE = REMARK.Substring(G3, G4 - G3 - 1).Replace("3.���:", "").Replace("\r", "");
                            MARK = REMARK.Substring(G4, G5 - G4 - 1).Replace("4.�Ƶ�:", "").Replace("\r", "");
                            EMAIL = REMARK.Substring(G5, GS.Length - G5).ToUpper().Replace("5.E-MAIL:", "").Replace("\r", "");

                            F29 = MARK;
                            F20 = EMAIL;
                        }

                    }
                }
            }
            string LineMaster =
                F01 + "|" +
                F02 + "|" +
                F03 + "|" +
                F04 + "|" +
                F05 + "|" +
                F06 + "|" +
                F07 + "|" +
                F08 + "|" +
                F09 + "|" +
                F10 + "|" +
                F11 + "|" +
                F12 + "|" +
                F13 + "|" +
                F14 + "|" +
                F15 + "|" +
                F16 + "|" +
                F17 + "|" +
                F18 + "|" +
                F19 + "|" +
                F20 + "|" +
                F21 + "|" +
                F22 + "|" +
                F23 + "|" +
                F24 + "|" +
                F25 + "|" +
                F26 + "|" +
                F27 + "|" +
                F28 + "|" +
                F29 + "|" +
                F30 + "|" +
                F31 + "|" +
                F32 + "|" +
                F33 + CrLf;

            sw.Write(LineMaster);



            //������
            //D01  ���ӥN��* C1 �T�w�� D
            string D01 = "D";
            //D02 �Ǹ�* C5
            string D02 = "";
            //D03 �q��s��* C40
            string D03 = "";
            //D04 �ӫ~�s�� C20
            string D04 = "";
            //D05 �ӫ~���X C20
            string D05 = "";
            //D06 �ӫ~�W��* C200
            string D06 = "";
            //D07 �ӫ~�W�� C100
            string D07 = "";
            //D08 ��� C6
            string D08 = "";
            //D09 ��� N14   ���� �p���I�H�U����
            string D09 = "";
            //D10 �ƶq* N13
            string D10 = "";
            //D11 ���|���B N14
            string D11 = "";
            //D12 �t�|���B* N14
            string D12 = "";
            //D13 ���d��  N13
            string D13 = "";
            //D14 �|�v�O*  C1  1: ���| 2: �s�|�v 3: �K�|
            string D14 = "3";
            //D15 ���Q�I�Ƨ馩���B  N13
            string D15 = "";
            //D16 ���ӳƵ� C100
            string D16 = "";
            //DEnd  ��Ƶ������̫�@�� ���� ORDER COUNT 
            string DEnd = "";



            string LineDetails =
               D01 + "|" +
               D02 + "|" +
               D03 + "|" +
               D04 + "|" +
               D05 + "|" +
               D06 + "|" +
               D07 + "|" +
               D08 + "|" +
               D09 + "|" +
               D10 + "|" +
               D11 + "|" +
               D12 + "|" +
               D13 + "|" +
               D14 + "|" +
               D15 + "|" +
               D16 + CrLf;
            if (OrderNo2 == "���J��")
            {

                System.Data.DataTable GG1 = GetREMARK2(OrderNumber);
                 if(GG1.Rows.Count > 0)
                {
         
                    dr = dtDetail.Rows[0];

                    ////������
                    ////D01  ���ӥN��* C1 �T�w�� D
                    D01 = "D";
                    ////D02 �Ǹ�* C5
                    D02 = (1).ToString();
                    ////D03 �q��s��* C40
                    BindData(dr, "BillNO", ref D03);
                    ////D04 �ӫ~�s�� C20
                    BindData(dr, "ProdID", ref D04);
                    ////D05 �ӫ~���X C20
                    PROD = D04;
                    ////D06 �ӫ~�W��* C200

                    D06 = ProdName;
                    //BindData(dr, D06 = "";
                    ////D07 �ӫ~�W�� C100
                    //BindData(dr, D07 = "";
                    ////D08 ��� C6
                    //BindData(dr, D08 = "";
                    ////D09 ��� N14   ���� �p���I�H�U����
                    D09 = PRICE;
                    D10 = QTY;

                    ////D11 ���|���B N14
                    BindData(dr, "Amount", ref D11, "");
                    ////D12 �t�|���B* N14
                    // BindData(dr, "MLAmount", D12);

                    Int32 x = Convert.ToInt32(Convert.ToDecimal(D11)) + Convert.ToInt32(dr["TaxAmt"]);

                    //Int32 x = Convert.ToInt16(D11) ;
                    D12 = Convert.ToString(x);
                    ////D13 ���d��  N13
                    D13 = "0";
                    ////D14 �|�v�O*  C1  1: ���| 2: �s�|�v 3: �K�|
                    //TaxType=0  ���|
                    //TaxType=1  �K�|
                    D14 = F06;
                    //BindData(dr, D14 = "3";
                    ////D15 ���Q�I�Ƨ馩���B  N13
                    D15 = "0";
                    //BindData(dr, D15 = "";
                    ////D16 ���ӳƵ� C100
                    D16 = "";
                    //BindData(dr, D16 = "";


                    LineDetails =
                       D01 + "|" +
                       D02 + "|" +
                       D03 + "|" +
                       D04 + "|" +
                       D05 + "|" +
                       D06 + "|" +
                       D07 + "|" +
                       D08 + "|" +
                       D09 + "|" +
                       D10 + "|" +
                       D11 + "|" +
                       D12 + "|" +
                       D13 + "|" +
                       D14 + "|" +
                       D15 + "|" +
                       D16 + CrLf;

                    sw.Write(LineDetails);
                }
            }
            else
            {
                for (int i = 0; i <= dtDetail.Rows.Count - 1; i++)
                {

                    dr = dtDetail.Rows[i];

                    ////������
                    ////D01  ���ӥN��* C1 �T�w�� D
                    D01 = "D";
                    ////D02 �Ǹ�* C5
                    D02 = (i + 1).ToString();
                    ////D03 �q��s��* C40
                    BindData(dr, "BillNO", ref D03);
                    ////D04 �ӫ~�s�� C20
                    BindData(dr, "ProdID", ref D04);
                    ////D05 �ӫ~���X C20
                    PROD = D04;
                    ////D06 �ӫ~�W��* C200
                    BindData(dr, "ProdName", ref D06);
                    D06 = D06.Replace("|", "");
                    //BindData(dr, D06 = "";
                    ////D07 �ӫ~�W�� C100
                    //BindData(dr, D07 = "";
                    ////D08 ��� C6
                    //BindData(dr, D08 = "";
                    ////D09 ��� N14   ���� �p���I�H�U����
                    BindData(dr, "Price", ref D09);
                    //BindData(dr, D09 = "";
                    ////D10 �ƶq* N13
                    BindData(dr, "Quantity", ref D10);
                    ////D11 ���|���B N14
                    BindData(dr, "Amount", ref D11, "");
                    ////D12 �t�|���B* N14
                    // BindData(dr, "MLAmount", D12);

                    Int32 x = Convert.ToInt32(Convert.ToDecimal(D11)) + Convert.ToInt32(dr["TaxAmt"]);

                    //Int32 x = Convert.ToInt16(D11) ;
                    D12 = Convert.ToString(x);
                    ////D13 ���d��  N13
                    D13 = "0";
                    ////D14 �|�v�O*  C1  1: ���| 2: �s�|�v 3: �K�|
                    //TaxType=0  ���|
                    //TaxType=1  �K�|
                    D14 = F06;
                    //BindData(dr, D14 = "3";
                    ////D15 ���Q�I�Ƨ馩���B  N13
                    D15 = "0";
                    //BindData(dr, D15 = "";
                    ////D16 ���ӳƵ� C100
                    D16 = "";
                    //BindData(dr, D16 = "";


                    LineDetails =
                       D01 + "|" +
                       D02 + "|" +
                       D03 + "|" +
                       D04 + "|" +
                       D05 + "|" +
                       D06 + "|" +
                       D07 + "|" +
                       D08 + "|" +
                       D09 + "|" +
                       D10 + "|" +
                       D11 + "|" +
                       D12 + "|" +
                       D13 + "|" +
                       D14 + "|" +
                       D15 + "|" +
                       D16 + CrLf;

                    sw.Write(LineDetails);
                }
            }


        }
        private void BindData(DataRow dr, string FromField, ref string ToField)
        {
            ToField = Convert.ToString(dr[FromField]);
            if (FromField == "ProdName")
            {
                System.Data.DataTable K1 = GetOITM(PROD);
                if (K1.Rows.Count > 0)
                {
                    ToField = K1.Rows[0][0].ToString();
                }
            }
        }
        private DataTable GetOITM(string ProdID)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();

            sb.Append("       select InvoProdName From comProduct A  Where A.ProdID = @ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private System.Data.DataTable GetCARD(string ORDERPIN)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  select SUBSTRING(CARD,13,4) from GB_POTATO  WHERE  ISNULL(CARD,'') <> '' AND ORDERPIN=@ORDERPIN ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ORDERPIN", ORDERPIN));



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
        private void BindData(DataRow dr, string FromField, ref string ToField, string IsInteger)
        {
            ToField = Convert.ToString(Convert.ToInt32(dr[FromField]));
        }
        private string ConvertDate(string sDate)
        {
            return sDate.Substring(0, 4) + "/" + sDate.Substring(4, 2) + "/" + sDate.Substring(6, 2);
        }
        public DataTable GetData(string Sql)
        {
            SqlConnection connection = new SqlConnection(strCn);


            SqlCommand command = new SqlCommand();
            command.Connection = connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(Sql);



            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();
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


        private void DELETEFILE2()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MAKE("1", "Y");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MAKE("3", "Y");
        }

        private DataTable GetREMARK(string FundBillNO, string DOCTYPE)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();

            sb.Append("  select Remark  from comBillAccounts where FundBillNO=@FundBillNO");
            if (DOCTYPE == "�P�h")
            {
                sb.Append(" and Flag=600  ");
            }
            else if (DOCTYPE == "����")
            {
                sb.Append(" and Flag=701  ");
            }
            else
            {
                sb.Append(" and Flag=500  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@FundBillNO", FundBillNO));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }
        private DataTable GetREMARK2(string FundBillNO)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();

            sb.Append(" 			   select REMARK from comBillAccounts where Flag IN  (595)  and FundBillNO =@FundBillNO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@FundBillNO", FundBillNO));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private DataTable GETCLASSID(string CLASSID)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT * FROM comCustClass where CLASSID IN ('011','013','014','015','018','020','021','022','026','027','028','029','030','031','032','033','034','036','037') AND CLASSID=@CLASSID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CLASSID", CLASSID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }
        private DataTable GetPAY(string OriginBillNo)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT FundBillNo FNO FROM ComFundSub WHERE OriginBillNo =@OriginBillNo AND OriginFlag =500");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@OriginBillNo", OriginBillNo));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }
        private void button8_Click(object sender, EventArgs e)
        {

            APS1CHOICE frm1 = new APS1CHOICE();
            frm1.CARDTYPE = "�q�l�o��";
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox4.Checked = true;
                c = frm1.q;

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Int32 iTotal = 0;

            int i = dataGridView1.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView1.SelectedRows[iRecs].Cells["���|���B"].Value);
            }

            textBox11.Text = iTotal.ToString("#,##0");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            MAKE("0", "N");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MAKE("1", "N");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            MAKE("3", "N");
        }

        private void button13_Click(object sender, EventArgs e)
        {

            if (isValidConnection(host, username, password) == true)
            {
                MessageBox.Show("�s�u���\");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            MAKE("3", "N");
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            MAKE("4", "N");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            MAKE("4", "Y");
        }



        private void dataGridView3_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {
                
                string BillNO = dataGridView3.SelectedRows[0].Cells["BillNO"].Value.ToString();
                string RowNO = dataGridView3.SelectedRows[0].Cells["RowNO"].Value.ToString();
                string Flag = dataGridView3.SelectedRows[0].Cells["Flag"].Value.ToString();
                string ���B = dataGridView3.SelectedRows[0].Cells["���B"].Value.ToString();
                if (���B == "0")
                {
                    UpdateBillNO(BillNO, Flag, RowNO);
                    MessageBox.Show("�w��s");
                    EXEC();
                }
            }
        }


        private void UpdateBillNO(string BillNO, string Flag, string RowNO)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE   stkBillSub SET IsGift=1 where BillNO=@BillNO AND Flag=@Flag AND RowNO=@RowNO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@Flag", Flag));
            command.Parameters.Add(new SqlParameter("@RowNO", RowNO));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
  
    }
}