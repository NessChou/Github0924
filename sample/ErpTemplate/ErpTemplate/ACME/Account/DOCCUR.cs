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
    public partial class DOCCUR : Form
    {
        public DOCCUR()
        {
            InitializeComponent();
        }


        private System.Data.DataTable DTOPOR(string ITEMCODE)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE,T1.QUANTITY �ƶq,CAST(T2.LINETOTAL/T1.TOTALFRGN AS DECIMAL(10,4)) �ײv FROM OPOR T0");
            sb.Append(" LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN PDN1 T2 ON (T1.docentry=T2.baseentry AND T1.linenum=T2.baseline)");
            sb.Append("  WHERE T1.ITEMCODE=@ITEMCODE  AND T1.TOTALFRGN <> 0 AND t2.basetype='22'  ");
            sb.Append(" ORDER BY T0.DOCDATE DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button1_Click(object sender, EventArgs e)
        {

        

            System.Data.DataTable dtCost = MakeTableCombine();
            System.Data.DataTable dt = OINV();
           
            DataRow dr = null;

           
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                dr["��f��date"] = dt.Rows[i]["��f��date"].ToString();
                dr["��f��#"] = dt.Rows[i]["��f��#"].ToString();
                dr["�P��q��#"] = dt.Rows[i]["�P��q��#"].ToString();
                dr["AR�o��"] = dt.Rows[i]["AR�o��"].ToString();
                dr["�Ȥ�s��"] = dt.Rows[i]["�Ȥ�s��"].ToString();
                dr["�Ȥ�X"] = dt.Rows[i]["�Ȥ�X"].ToString();
                dr["�Ȥ�W��"] = dt.Rows[i]["�Ȥ�W��"].ToString();
                dr["���O"] = dt.Rows[i]["���O"].ToString();
                decimal ���|��l���B = Convert.ToDecimal(dt.Rows[i]["���|��l���B"].ToString());
                dr["���|��l���B"] = ���|��l���B.ToString("#,##0");
                decimal �ײv = Convert.ToDecimal(dt.Rows[i]["�ײv"].ToString());
                dr["�ײv"] = �ײv.ToString("#,##0.0000");
                decimal �b�����B = Convert.ToDecimal(dt.Rows[i]["�b�����B"].ToString());
                dr["�b�����B"] = �b�����B.ToString("#,##0");
                decimal �P�f���� = Convert.ToDecimal(dt.Rows[i]["�P�f����"].ToString());
                dr["�P�f����"] = �P�f����.ToString("#,##0.0000");
                decimal �P��q��ײv = Convert.ToDecimal(dt.Rows[i]["�q��ײv"].ToString());
                dr["�P��q��ײv"] = �P��q��ײv.ToString("#,##0.0000");
                string ITEMCODE = dt.Rows[i]["�~�W"].ToString();
                dr["�~�W"] = ITEMCODE;
                decimal QTY = Convert.ToDecimal(dt.Rows[i]["�ƶq"].ToString());
                dr["�ƶq"] = QTY.ToString("#,##0");
                DataTable dt1 = DTOPOR(ITEMCODE);
                decimal QUANTITY = 0;
                decimal FINAL = 0;
                decimal DD = 0;
                for (int j = 0; j <= dt1.Rows.Count - 1; i++)
                {
                    decimal aa = Convert.ToDecimal(dt1.Rows[j]["�ײv"].ToString());
                    decimal bb = Convert.ToDecimal(dt1.Rows[j]["�ƶq"].ToString());
              

                    QUANTITY += bb;
                    decimal F1 = QTY - QUANTITY;
                    if (F1 < 0)
                    {
                        DD = QTY - (QUANTITY - bb);
                    }
                    if (F1 >= 0)
                    {
                        FINAL += aa * bb;
                    }
                    else
                    {
                        FINAL += aa * (DD);
                        break;
                    }
                }

                decimal ���ʶײv = FINAL / QTY;
                dr["���ʶײv"] = (���ʶײv).ToString("#,##0.0000");
                decimal �קI�l�q = (�ײv - ���ʶײv) * ���|��l���B;
                dr["�קI�l�q"] = (�קI�l�q).ToString("#,##0.0000");

                decimal �קI�l�q2 = (�ײv - �P��q��ײv) * ���|��l���B;
                dr["�קI�l�q2"] = (�קI�l�q2).ToString("#,##0.0000");
                dtCost.Rows.Add(dr);


            }
            dataGridView1.DataSource = dtCost;

            for (int i = 8; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

        }
        private System.Data.DataTable OINV()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  Convert(varchar(8),T7.docdate,112) ��f��date,T7.DOCENTRY ��f��#,T8.DOCENTRY �P��q��#,T0.DOCENTRY AR�o��");
            sb.Append(" ,SUBSTRING(T11.GROUPNAME,4,5) �Ȥ�s��,T0.CARDCODE �Ȥ�X,T0.CARDNAME �Ȥ�W��,T9.DOCCUR ���O,T8.TOTALFRGN ���|��l���B,T1.LINETOTAL �b�����B");
            sb.Append(" ,CAST(T1.LINETOTAL/T8.TOTALFRGN AS DECIMAL(10,4)) �ײv,T1.GrossBuyPr �P�f����,T1.QUANTITY �ƶq,T1.PRICE ���,T1.ITEMCODE �~�W,ISNULL(T8.RATE,0) �q��ײv FROM OINV T0  ");
            sb.Append(" LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append(" LEFT JOIN DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append(" LEFT JOIN RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append(" LEFT JOIN ORDR T9 ON (T8.docentry=T9.docentry )");
            sb.Append(" LEFT JOIN OCRD T10 ON (T0.CARDCODE=T10.CARDCODE)");
            sb.Append(" LEFT JOIN OCRG T11 ON (T10.GROUPCODE=T11.GROUPCODE)");
            sb.Append(" where t1.basetype='15' AND T8.GtotalFC <> 0 AND T9.DOCCUR='USD' ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append("   and Convert(varchar(8),T7.docdate,112) between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'"); 
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {
                sb.Append("   and T7.DOCENTRY between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "'");
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {
                sb.Append("   and T9.DOCENTRY between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "'");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("��f��date", typeof(string));
            dt.Columns.Add("��f��#", typeof(string));
            dt.Columns.Add("�P��q��#", typeof(string));
            dt.Columns.Add("AR�o��", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�X", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("���O", typeof(string));
            dt.Columns.Add("���|��l���B", typeof(string));
            dt.Columns.Add("�ײv", typeof(string));
            dt.Columns.Add("�b�����B", typeof(string));
            dt.Columns.Add("�P�f����", typeof(string));
            dt.Columns.Add("���ʶײv", typeof(string));
            dt.Columns.Add("�קI�l�q", typeof(string));
            dt.Columns.Add("�P��q��ײv", typeof(string));
            dt.Columns.Add("�קI�l�q2", typeof(string));
            dt.Columns.Add("�~�W", typeof(string));
            dt.Columns.Add("�ƶq", typeof(string));
            return dt;
        }

        private void DOCCUR_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}