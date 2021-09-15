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
namespace ACME
{
    public partial class POTATOAR : Form
    {
        private SerialPort comport = new SerialPort();
        StringBuilder sb = new StringBuilder();
        string MESS = "";
        public POTATOAR()
        {
            InitializeComponent();
        }

        private void POTATOAR_Load(object sender, EventArgs e)
        {

            comboBox3.Text = "���f���";
            comboBox4.Text = "�ɧ�";
            btnPrintTest.Enabled = false;
            if (comport.IsOpen) comport.Close();
            else
            {
                //�]�w��
                comport.BaudRate = 9600;
                comport.DataBits = 8;
                comport.StopBits = StopBits.One;
                comport.Parity = Parity.None;
                comport.PortName = "COM1";
                try
                {
                    comport.Open();
                }
                catch 
                {
                    //MessageBox.Show(ex.Message);
                   // return;
                }
            }

            if (comport.IsOpen)
            {
                MessageBox.Show("�o�����w���\�s��");
                btnPrintTest.Enabled = true;
             
            }

            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();


            comboBox2.Items.Clear();

            System.Data.DataTable dt3 = GetOrderData3V();

            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }

            comboBox2.Items.Add("����");
        }
        public static void Order(SerialPort printer, byte[] command)
        {
            printer.Write(command, 0, command.Length);
        }
        private void btnPrintTest_Click(object sender, EventArgs e)
        {
            int f1 = 0;
            for (int j = 0; j <= dataGridView1.SelectedRows.Count - 1; j++)
            {
                string F = dataGridView1.Rows[j].Cells["�o�����X"].Value.ToString();
  
                if (String.IsNullOrEmpty(F))
                {
                    f1 = 1;
                }
            }

            if (f1 == 1)
            {
                MessageBox.Show("�z���ťժ��o�����X");
                return;
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
                    string INV = dataGridView1.SelectedRows[j].Cells["�νs"].Value.ToString();
                    string DOC = dataGridView1.SelectedRows[j].Cells["���f���"].Value.ToString();
                    bool openMoneyBox_BeforePrinting = true;
                    bool openMoneyBox_AfterPrinting = true;

                    comport.Encoding = Encoding.Default;

                    // comport.Order(Command.ResetPrinter);    //��l�L���
                    comport.Write(Command.ResetPrinter, 0, Command.ResetPrinter.Length);
                    // comport.Order(Command.StubAndReceiver); //�����p�P�s���p�P�ɦC�L
                    comport.Write(Command.StubAndReceiver, 0, Command.StubAndReceiver.Length);
                    // comport.Order(Command.MoveLines(4));    //���L�e4��
                    //  comport.Write(Command.MoveLines(4), 0, Command.MoveLines(4).Length);

                    if (openMoneyBox_BeforePrinting)
                        //comport.Order(Command.OpenMoneyBox1);
                        comport.Write(Command.OpenMoneyBox1, 0, Command.OpenMoneyBox1.Length);
                    System.Data.DataTable T1 = GetOrderData31(F2);
                    System.Data.DataTable T3 = GetOrderData32(F2);
                    //QTY,T0.PRICE,AMOUNT,INVNAME
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
               
                    //comport.WriteLine("�զX�y�� 1 x 5600  5,600");
                    //comport.WriteLine("��Ʈw   2 x 5600 11,200");
                    if (T3.Rows.Count > 0)
                    {
                        for (int i = 0; i <= T1.Rows.Count - 1; i++)
                        {
                            string INVNAME = T1.Rows[i]["INVNAME"].ToString();
                            string ITEMCODE = T1.Rows[i]["ITEMCODE"].ToString();

                            System.Data.DataTable L11 = GetOITM(ITEMCODE);
                            DataRow dv1 = L11.Rows[0];
                            string ITEMTYPE = dv1["ITEMTYPE"].ToString();
              
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

                            int TPRICE = Convert.ToInt16(T1.Rows[i]["PRICE"].ToString());

                            string PRICE = TPRICE.ToString("#,##0");
                            int PRICEINT = PRICE.Length;

                            if (PRICEINT == 3)
                            {
                                PRICE = "  " + PRICE;
                            }


                            int TAMOUNT = Convert.ToInt16(T1.Rows[i]["AMOUNT"].ToString());
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

                            if (ITEMTYPE == "B")
                            {
                               // comport.WriteLine(ITEMCODE + " " + 1 + "  X" + PRICE + AMOUNT + "NX");
                                comport.WriteLine(ITEMCODE + "   " + 1 + "       "  + AMOUNT + "NX");
                                System.Data.DataTable L1 = GetBOM(ITEMCODE);
                                for (int H = 0; H <= L1.Rows.Count - 1; H++)
                                {
                                    string ITEM = L1.Rows[H]["INVNAME"].ToString();
                                    string QUANTITY = L1.Rows[H]["QTY"].ToString();
                                    int QUANTITYT = QUANTITY.Length;
                                    if (QUANTITYT == 1)
                                    {
                                        QUANTITY = "  " + QUANTITY;
                                    }
                                    else if (QUANTITYT == 2)
                                    {
                                        QUANTITY = " " + QUANTITY;
                                    }
                                    comport.WriteLine(ITEM + " " + QUANTITY);
                                }
                            }
                            else
                            {
                               // comport.WriteLine(INVNAME + " " + QTY + "  X" + PRICE + AMOUNT + "NX");
                                comport.WriteLine(INVNAME + QTY + "     " + AMOUNT + "NX");
                            }

                        }


                        int T���B = Convert.ToInt16(T3.Rows[0]["���B"].ToString());
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
                        //
                        int T�B�O = Convert.ToInt16(T3.Rows[0]["�B�O"].ToString());
                        string �B�O = T�B�O.ToString("#,##0");
                        int �B�OT = �B�O.Length;
                        if (�B�OT == 6)
                        {
                            �B�O = "           " + �B�O;
                        }
                        if (�B�OT == 5)
                        {
                            �B�O = "            " + �B�O;
                        }
                        if (�B�OT == 3)
                        {
                            �B�O = "              " + �B�O;
                        }
                        if (�B�OT == 1)
                        {
                            �B�O = "                " + �B�O;
                        }

                        int T�`�p = Convert.ToInt16(T3.Rows[0]["�`�p"].ToString());
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
                        comport.WriteLine("�p�p:" + ���B + "NX");
                        if (�B�O != "0")
                        {
                            comport.WriteLine("�B�O:" + �B�O + "TX");
                        }
                        comport.WriteLine("========================");
                        comport.WriteLine("�`�p:" + �`�p);

                        int F = 0;
                        int F3 = 0;
                        if (T�B�O != 0)
                        {
                            comport.Write(Command.MoveLines(1), 0, Command.MoveLines(1).Length);
                            F = Convert.ToInt16(T�B�O / 1.05);
                            F3 = T�B�O - F;

                            string ���|���B = F.ToString("#,##0");
                            int ���|���BT = ���|���B.Length;
                            if (���|���BT == 7)
                            {
                                ���|���B = "        " + ���|���B;
                            }
                            if (���|���BT == 6)
                            {
                                ���|���B = "         " + ���|���B;
                            }
                            if (���|���BT == 5)
                            {
                                ���|���B = "          " + ���|���B;
                            }
                            if (���|���BT == 3)
                            {
                                ���|���B = "            " + ���|���B;
                            }

                            string �|�� = F3.ToString("#,##0");
                            int �|��T = �|��.Length;
                            if (�|��T == 4)
                            {
                                �|�� = "               " + �|��;
                            }
                            if (�|��T == 3)
                            {
                                �|�� = "                " + �|��;
                            }
                            if (�|��T == 2)
                            {
                                �|�� = "                 " + �|��;
                            }
                            if (�|��T == 1)
                            {
                                �|�� = "                  " + �|��;
                            }

                            comport.WriteLine("���|���B " + ���|���B);
                            comport.WriteLine("�|�B " + �|��);
                            comport.WriteLine("�K�|���B " + �K�|���B);


                        }

                        comport.Write(Command.MoveLines(1), 0, Command.MoveLines(1).Length);
                        comport.WriteLine("PO# " + F2);
                  
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
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            EXEC();

            ////DataGridView1.Rows[2].Frozen = true;

            //MessageBox.Show(DATAGGR
        }
        private void EXEC()
        {
            System.Data.DataTable dt = GetOrderData3("2");
            if (dt.Rows.Count == 0)
            {
                System.Data.DataTable dt2 = GetOrderData3("1");
                dataGridView1.DataSource = dt2;
            }
            else
            {

                dataGridView1.DataSource = dt;
                dataGridView1.Columns["�o�����X"].ReadOnly = false;


                for (int j = 0; j <= dataGridView1.Rows.Count - 1; j++)
                {
                    string F = dataGridView1.Rows[j].Cells["AFEE"].Value.ToString();

                    if (F.Trim() == "True")
                    {
                        dataGridView1.Rows[j].Cells["�o�����X"].ReadOnly = true;
                    }
                }



                DataRow row;
                //�[�J�@���X�p
                Int32[] Total = new Int32[dt.Columns.Count - 1];

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    for (int j = 2; j <= 8 ; j++)
                    {
                        try
                        {
                            Total[j - 1] += Convert.ToInt32(dt.Rows[i][j]);
                        }
                        catch
                        {
                            Total[j - 1] += 0;
                        }

                        //if (j == 3 || j == 5)
                        //{
                        //    Total[j - 1] = 0;
                        //}
                    }
                }



                row = dt.NewRow();

                row[1] = "�X�p";
                for (int j = 2; j <= 8; j++)
                {
                    row[j] = Total[j - 1];

                    //if (j == 3 || j == 5)
                    //{
                    //    row[j] = 0;
                    //}

                }
                dt.Rows.Add(row);
            }
        }
        private System.Data.DataTable GetOrderData3(string A)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                                        SELECT T0.ID,PAYMAN �I�ڤH,QTY1 ����,QTY2 �b��");
            sb.Append("                                                  ,U_IN_BSAMN ���|���B,U_IN_BSTAX �|�B,U_IN_BSAMT �K�|���B,SHIPFEE �B�O,Amount �`�p,");
            sb.Append(" (SELECT MAX(DELREMARK) A  FROM GB_FRIEND WHERE DOCID=T0.ID)  ���f���,");
            sb.Append("                                                  TransMark ����覡,UNIT �νs,T0.SHIPDATE �o�����X,CASE ISNULL(AFEE,'') WHEN '' THEN 'False' else AFEE end AFEE FROM dbo.GB_POTATO T0   ");
            sb.Append("                                           WHERE 1=1   and TransMark <> 'FOC'");
            sb.Append("          AND T0.ID IN (SELECT DISTINCT DOCID FROM GB_FRIEND  WHERE DELREMARK BETWEEN @CreateDate AND @CreateDate1) ");
            if (comboBox1.Text == "�w�}�ߵo��")
            {
                sb.Append("  AND ISNULL(T0.SHIPDATE,'') <> '' ");
            
            }
            if (comboBox1.Text == "���}�ߵo��")
            {
                sb.Append("  AND ISNULL(T0.SHIPDATE,'') = '' ");

            }
            if (comboBox2.Text != "" && comboBox2.Text != "����")
            {
                sb.Append(" AND TransMark=@TransMark ");

            }
            if (A == "1")
            {
                sb.Append(" AND 1 <> 2 ");
            }


            if (comboBox3.Text == "PO")
            {
                sb.Append(" ORDER BY ID");
            }
            else if (comboBox3.Text == "�I�ڤH")
            {
                sb.Append(" ORDER BY PAYMAN");
            }
            else if (comboBox3.Text == "�����ƶq")
            {
                sb.Append(" ORDER BY QTY1");
            }
            else if (comboBox3.Text == "�������")
            {
                sb.Append(" ORDER BY Qty1P");
            }
            else if (comboBox3.Text == "�b���ƶq")
            {
                sb.Append(" ORDER BY QTY2");
            }
            else if (comboBox3.Text == "�b�����")
            {
                sb.Append(" ORDER BY Qty2P");
            }
            else if (comboBox3.Text == "���|���B")
            {
                sb.Append(" ORDER BY U_IN_BSAMN");
            }
            else if (comboBox3.Text == "�|�B")
            {
                sb.Append(" ORDER BY U_IN_BSTAX");
            }
            else if (comboBox3.Text == "�K�|���B")
            {
                sb.Append(" ORDER BY U_IN_BSAMT");
            }
            else if (comboBox3.Text == "�B�O")
            {
                sb.Append(" ORDER BY SHIPFEE");
            }
            else if (comboBox3.Text == "�`�p")
            {
                sb.Append(" ORDER BY Amount");
            }
            else if (comboBox3.Text == "���f���")
            {
                sb.Append(" ORDER BY (SELECT MAX(DELREMARK) A  FROM GB_FRIEND WHERE DOCID=T0.ID)");
            }
            else if (comboBox3.Text == "����覡")
            {
                sb.Append(" ORDER BY TransMark");
            }
            else if (comboBox3.Text == "�Τ@�s��")
            {
                sb.Append(" ORDER BY UNIT");
            }
            else if (comboBox3.Text == "�o�����X")
            {
                sb.Append(" ORDER BY T0.SHIPDATE");
            }
            else if (comboBox3.Text == "�o���}�ߧ���")
            {
                sb.Append(" ORDER BY CASE ISNULL(AFEE,'') WHEN '' THEN 'False' else AFEE end");
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
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CreateDate1", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@TransMark", comboBox2.Text));
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
        private System.Data.DataTable GetOrderData31(string ID)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT QTY,T0.PRICE,AMOUNT,INVNAME,T0.ITEMCODE,ITEMTYPE FROM dbo.GB_POTATO2 T0");
            sb.Append(" LEFT JOIN GB_OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append(" WHERE T0.ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
    

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

        private System.Data.DataTable GetOrderData32(string ID)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PotatoWg ���B,T0.SHIPFEE �B�O,AMOUNT �`�p FROM dbo.GB_POTATO T0");
            sb.Append(" WHERE T0.ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));


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
        private System.Data.DataTable GetBOM(string FATHER)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT QTY,INVNAME FROM GB_BOM T0  LEFT JOIN GB_OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)  WHERE FATHER=@FATHER ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@FATHER", FATHER));


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
        private System.Data.DataTable GetOrderData3V()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                            SELECT PARAM_NO FROM PARAMS WHERE PARAM_KIND='POTATOTYPE3'");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


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
        private void button2_Click(object sender, EventArgs e)
        {


            //SELECTIDT2();

            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("�п�ܭn�������C");
                return;
            }
           
                for (int j = dataGridView1.SelectedRows.Count - 1; j >= 0; j--)
                {
                    //���f���
                    string F = dataGridView1.SelectedRows[j].Cells["�o�����X"].Value.ToString();
                    string F2 = dataGridView1.SelectedRows[j].Cells["ID"].Value.ToString();
                    string ���f��� = dataGridView1.SelectedRows[j].Cells["���f���"].Value.ToString();
                    if (String.IsNullOrEmpty(F))
                    {
                        System.Data.DataTable T1 = GetOrderData4();
                        if (T1.Rows.Count > 0)
                        {
                            string ft = T1.Rows[0][0].ToString();
                            string ID = T1.Rows[0][1].ToString();
                            string ZQ = T1.Rows[0][2].ToString().Trim();
                            dataGridView1.SelectedRows[j].Cells["�o�����X"].Value = ZQ + ft;

                            if (String.IsNullOrEmpty(ZQ))
                            {
                                MessageBox.Show("�S���r�y");
                                return;

                            }
                            UpdateID(ft, ID);
                            UpdateID2(ZQ+ft, F2);
                        }
                    }
                }
            
        }

        private System.Data.DataTable GetOrderData4()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT TOP 1 CASE ISNULL(U_BSRN3,0) WHEN 0 THEN U_BSRN1 ELSE U_BSRN3+1 END INV,ID,U_BSTRK FROM dbo.GB_INVTRACK");
            sb.Append("  WHERE  @CreateDate BETWEEN U_BSYNM AND U_BSYEM ");
            sb.Append(" AND U_BSRN2 <> ISNULL(U_BSRN3,0)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text.Substring(0, 8)));


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
        private System.Data.DataTable GetOrderData4T()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT TOP 1 CASE ISNULL(U_BSRN3,0) WHEN 0 THEN U_BSRN1 ELSE U_BSRN3+1 END INV,ID FROM " + "AR" + fmLogin.LoginID.ToString() + "");
            sb.Append("  WHERE Convert(varchar(11),GETDATE(),112) BETWEEN @CreateDate AND @CreateDate1");
            sb.Append(" AND U_BSRN2 <> ISNULL(U_BSRN3,0)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CreateDate1", textBox2.Text));

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
        //private System.Data.DataTable GetOrderData5(string U_BSRN3)
        //{

        //    SqlConnection connection = globals.Connection;

        //    StringBuilder sb = new StringBuilder();

        //    sb.Append(" SELECT TOP 1 CASE ISNULL(U_BSRN3,0) WHEN 0 THEN U_BSRN1 ELSE U_BSRN3+1 END INV,ID FROM dbo.GB_INVTRACK");
        //    sb.Append("  WHERE Convert(varchar(11),GETDATE(),112) BETWEEN U_BSYMM AND U_BSYEM");
        //    sb.Append(" AND U_BSRN2 <> ISNULL(U_BSRN3,0) AND U_BSRN2 <> ISNULL(@U_BSRN3,0) ");

        //    SqlCommand command = new SqlCommand(sb.ToString(), connection);
        //    command.CommandType = CommandType.Text;
        //    command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text));
        //    command.Parameters.Add(new SqlParameter("@CreateDate1", textBox2.Text));
        //    command.Parameters.Add(new SqlParameter("@U_BSRN3", U_BSRN3));
        //    SqlDataAdapter da = new SqlDataAdapter(command);

        //    DataSet ds = new DataSet();
        //    try
        //    {
        //        connection.Open();
        //        da.Fill(ds, "rma_PackingListM");
        //    }
        //    finally
        //    {
        //        connection.Close();
        //    }

        //    return ds.Tables[0];

        //}
        private void UpdateID(string U_BSRN3, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE  GB_INVTRACK SET U_BSRN3=@U_BSRN3 WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@U_BSRN3", U_BSRN3));
            command.Parameters.Add(new SqlParameter("@ID", ID));

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

        //private void UpdateIDT(string U_BSRN3, string ID)
        //{
        //    SqlConnection connection = globals.Connection;
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append(" UPDATE  " + "AR" + fmLogin.LoginID.ToString() + " SET U_BSRN3=@U_BSRN3 WHERE ID=@ID");

        //    SqlCommand command = new SqlCommand(sb.ToString(), connection);
        //    command.CommandType = CommandType.Text;
        //    SqlDataAdapter da = new SqlDataAdapter(command);
        //    command.Parameters.Add(new SqlParameter("@U_BSRN3", U_BSRN3));
        //    command.Parameters.Add(new SqlParameter("@ID", ID));

        //    try
        //    {

        //        try
        //        {
        //            connection.Open();
        //            command.ExecuteNonQuery();
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message);
        //        }
        //    }
        //    finally
        //    {
        //        connection.Close();
        //    }


        //}
        private void SELECTIDT()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            //sb.Append(" DROP TABLE " + "AR" + fmLogin.LoginID.ToString() + "");
            sb.Append(" SELECT * INTO " + "AR" + fmLogin.LoginID.ToString() + " FROM dbo.GB_INVTRACK");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
           // command.Parameters.Add(new SqlParameter("@AA", "AR"+fmLogin.LoginID.ToString()));

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
        private void SELECTIDT2()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" DROP TABLE " + "AR" + fmLogin.LoginID.ToString() + "");
            //sb.Append(" SELECT * INTO " + "AR" + fmLogin.LoginID.ToString() + " FROM dbo.GB_INVTRACK");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            // command.Parameters.Add(new SqlParameter("@AA", "AR"+fmLogin.LoginID.ToString()));

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
        private void UpdateID2(string SHIPDATE, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE  GB_POTATO SET SHIPDATE=@SHIPDATE WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPDATE", SHIPDATE));
            command.Parameters.Add(new SqlParameter("@ID", ID));

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
        private void UpdateID3(string AFEE, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE  GB_POTATO SET AFEE=@AFEE WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@AFEE", AFEE));
            command.Parameters.Add(new SqlParameter("@ID", ID));

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
        private void button3_Click(object sender, EventArgs e)
        {
            for (int j = 0; j <= dataGridView1.Rows.Count - 1; j++)
            {
                string F = dataGridView1.Rows[j].Cells["�o�����X"].Value.ToString();

                string F2 = dataGridView1.Rows[j].Cells["ID"].Value.ToString();
                string AFEE = dataGridView1.Rows[j].Cells["AFEE"].Value.ToString();
                UpdateID3(AFEE, F2);
            }

            for (int j = 0; j <= dataGridView1.Rows.Count - 1; j++)
            {
                string F = dataGridView1.Rows[j].Cells["�o�����X"].Value.ToString();

                string F2 = dataGridView1.Rows[j].Cells["ID"].Value.ToString();
                string AFEE = dataGridView1.Rows[j].Cells["AFEE"].Value.ToString();
                if (AFEE != "True")
                {
                    UpdateID2(F, F2);
                }
            }
            MessageBox.Show("�ק令�\");
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
        private System.Data.DataTable GetOITM(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT INVNAME,ITEMTYPE,ITEMOI FROM GB_OITM WHERE ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));


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
                string INV = dataGridView1.SelectedRows[0].Cells["�νs"].Value.ToString();
                string DOC = dataGridView1.SelectedRows[0].Cells["���f���"].Value.ToString();


  

          
                System.Data.DataTable T1 = GetOrderData31(F2);
                System.Data.DataTable T3 = GetOrderData32(F2);

                string DOCDATE = DOC.Substring(0, 4) + "/" + DOC.Substring(4, 2) + "/" + DOC.Substring(6, 2);

                Zr.WriteLine("���׹�~�ѥ��������q" + System.Environment.NewLine);
                 Zr.WriteLine("��~�H�νs: 22468373" + System.Environment.NewLine);
                 Zr.WriteLine("�x�_������Ϸs��G��" + System.Environment.NewLine);
                 Zr.WriteLine("257��5�Ӥ�3 TEL:87922800" + System.Environment.NewLine);
                 Zr.WriteLine("POS# ARMAS-001" + System.Environment.NewLine);
                 Zr.WriteLine(DOCDATE + System.Environment.NewLine);
                if (!String.IsNullOrEmpty(INV))
                {
                     Zr.WriteLine("�Τ@�s��: " + System.Environment.NewLine);
                 
                }
                 Zr.WriteLine("------------------------" + System.Environment.NewLine);
                if (T3.Rows.Count > 0)
                {
                    for (int i = 0; i <= T1.Rows.Count - 1; i++)
                    {
                        string INVNAME = T1.Rows[i]["INVNAME"].ToString();
                        string ITEMCODE = T1.Rows[i]["ITEMCODE"].ToString();
                        System.Data.DataTable L11 = GetOITM(ITEMCODE);
                        DataRow dv1 = L11.Rows[0];
                        string ITEMTYPE = dv1["ITEMTYPE"].ToString();
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
                        int TPRICE = Convert.ToInt16(T1.Rows[i]["PRICE"].ToString());

                        string PRICE = TPRICE.ToString("#,##0");
                        int PRICEINT = PRICE.Length;

                        if (PRICEINT == 3)
                        {
                            PRICE = "  " + PRICE;
                        }
                       // string A1 = T1.Rows[i]["AMOUNT"].ToString();

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

                        if (ITEMTYPE == "B")
                        {
                          
                           //  Zr.WriteLine(ITEMCODE + " " + 1 + "  X" + PRICE + AMOUNT + "NX" + System.Environment.NewLine);
                            Zr.WriteLine(ITEMCODE + "   " + 1 + "       "  + AMOUNT + "NX" + System.Environment.NewLine);
                            System.Data.DataTable L1 = GetBOM(ITEMCODE);
                            for (int H = 0; H <= L1.Rows.Count - 1; H++)
                            {
                                string ITEM = L1.Rows[H]["INVNAME"].ToString();
                                string QUANTITY = L1.Rows[H]["QTY"].ToString();
                                int QUANTITYT = QUANTITY.Length;
                                if (QUANTITYT == 1)
                                {
                                    QUANTITY = "  " + QUANTITY;
                                }
                                else if (QUANTITYT == 2)
                                {
                                    QUANTITY = " " + QUANTITY;
                                }

                                 Zr.WriteLine(ITEM + " " + QUANTITY + System.Environment.NewLine);
                            }
                        }
                        else
                        {
                           //  Zr.WriteLine(INVNAME + " " + QTY + "  X" + PRICE + AMOUNT + "NX" + System.Environment.NewLine);
                            Zr.WriteLine(INVNAME + QTY + "     " + AMOUNT + "NX" + System.Environment.NewLine);
                        }

                    }


                    int T���B = Convert.ToInt16(T3.Rows[0]["���B"].ToString());
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

                    int T�B�O = Convert.ToInt16(T3.Rows[0]["�B�O"].ToString());
                    string �B�O = T�B�O.ToString("#,##0");
                    int �B�OT = �B�O.Length;
                    if (�B�OT == 6)
                    {
                        �B�O = "           " + �B�O;
                    }
                    if (�B�OT == 5)
                    {
                        �B�O = "            " + �B�O;
                    }
                    if (�B�OT == 3)
                    {
                        �B�O = "              " + �B�O;
                    }
                    if (�B�OT == 1)
                    {
                        �B�O = "                " + �B�O;
                    }

                    int T�`�p = Convert.ToInt16(T3.Rows[0]["�`�p"].ToString());
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
                     Zr.WriteLine("�p�p:" + ���B + "NX" + System.Environment.NewLine);

                    if (�B�O != "0")
                    {
                         Zr.WriteLine("�B�O:" + �B�O + "TX" + System.Environment.NewLine);
              
                    }
                     Zr.WriteLine("========================" + System.Environment.NewLine);
                     Zr.WriteLine("�`�p:" + �`�p + System.Environment.NewLine);

                    int F = 0;
                    int F3 = 0;
                    if (T�B�O != 0)
                    {
                        comport.Write(Command.MoveLines(1), 0, Command.MoveLines(1).Length);
                        F = Convert.ToInt16(T�B�O / 1.05);
                        F3 = T�B�O - F;

                        string ���|���B = F.ToString("#,##0");
                        int ���|���BT = ���|���B.Length;
                        if (���|���BT == 7)
                        {
                            ���|���B = "        " + ���|���B;
                        }
                        if (���|���BT == 6)
                        {
                            ���|���B = "         " + ���|���B;
                        }
                        if (���|���BT == 5)
                        {
                            ���|���B = "          " + ���|���B;
                        }
                        if (���|���BT == 3)
                        {
                            ���|���B = "            " + ���|���B;
                        }

                        string �|�� = F3.ToString("#,##0");
                        int �|��T = �|��.Length;
                        if (�|��T == 4)
                        {
                            �|�� = "               " + �|��;
                        }
                        if (�|��T == 3)
                        {
                            �|�� = "                " + �|��;
                        }
                        if (�|��T == 2)
                        {
                            �|�� = "                 " + �|��;
                        }
                        if (�|��T == 1)
                        {
                            �|�� = "                  " + �|��;
                        }
                         Zr.WriteLine("���|���B " + ���|���B + System.Environment.NewLine);
                         Zr.WriteLine("�|�B " + �|�� + System.Environment.NewLine);
                         Zr.WriteLine("�K�|���B " + �K�|���B + System.Environment.NewLine);
                     


                    }
                     Zr.WriteLine("" + System.Environment.NewLine);
                     Zr.WriteLine("PO# " + F2 + System.Environment.NewLine);

                     Zfs.Flush();
                     Zr.Close();
                     System.Diagnostics.Process.Start(ZfileName);
                
            }
        }

    


    }
}