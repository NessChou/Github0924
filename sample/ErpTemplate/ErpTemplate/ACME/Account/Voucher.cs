using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.IO;


namespace ACME
{
    public partial class Voucher : Form
    {
        DataTable dt = null;
        string ff;
 
        public Voucher()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
                object[] LookupValues = GetMenu.GetMenuListU();

                if (LookupValues != null)
                {
                    cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                    cardNameTextBox.Text = Convert.ToString(LookupValues[1]);

                }
         
            
        }

        private void button47_Click(object sender, EventArgs e)
        {
            try
            {
                if (cardNameTextBox.Text == "")
                {
                    MessageBox.Show("請輸入客戶");
                }
                else
                {
                    string DuplicateKey = "";
                    string 單號;
                    System.Data.DataTable dtemp5 = GetVoucher();
                    System.Data.DataTable dtCostDD = MakeTableFcstWeek();
                    DataRow drtemp5 = null;
                    for (int i = 0; i <= dtemp5.Rows.Count - 1; i++)
                    {
                        drtemp5 = dtCostDD.NewRow();
                        單號 = dtemp5.Rows[i]["ID"].ToString();
                        drtemp5["客戶編號"] = dtemp5.Rows[i]["客戶編號"].ToString();
                        drtemp5["客戶名稱"] = dtemp5.Rows[i]["客戶名稱"].ToString();

                        drtemp5["科目"] = dtemp5.Rows[i]["科目"].ToString();
                        drtemp5["名稱"] = dtemp5.Rows[i]["名稱"].ToString();
                       
                        drtemp5["借方"] = dtemp5.Rows[i]["借方"].ToString();

                        drtemp5["營業稅"] = 0;
                        if (單號 != DuplicateKey)
                        {
                            drtemp5["ID"] = 單號;
                            drtemp5["TERM"] = dtemp5.Rows[i]["TERM"].ToString();
                            drtemp5["部門"] = dtemp5.Rows[i]["部門"].ToString();
                            drtemp5["過帳日期"] = dtemp5.Rows[i]["過帳日期"].ToString();
                            drtemp5["營業稅"] = Convert.ToInt32(dtemp5.Rows[i]["營業稅"].ToString());
                            drtemp5["起始日"] = dtemp5.Rows[i]["起始日"].ToString();
                            drtemp5["結束日"] = dtemp5.Rows[i]["結束日"].ToString();
                            drtemp5["發票日期"] = dtemp5.Rows[i]["發票日期"].ToString();
                            string ff = dtemp5.Rows[i]["到期日"].ToString();
                            if (!String.IsNullOrEmpty(ff))
                            {
                                drtemp5["DUEDATE"] = ff.Substring(0, 4).ToString() + ff.Substring(5, 2).ToString() + ff.Substring(8, 2).ToString();
                            }

                            drtemp5["到期日"] = ff;
                            
                            drtemp5["單號"] = dtemp5.Rows[i]["單號"].ToString();
                         
                            string fh = dtemp5.Rows[i]["發票"].ToString();
                            int g = fh.IndexOf("__");
                              if (g == -1)
                              {
                                  drtemp5["發票"] = fh;
                              }
                    
                            else
                            {
                                System.Data.DataTable fg = GetVoucher2(單號);
                                StringBuilder sb = new StringBuilder();
                                for (int p = 0; p <= fg.Rows.Count - 1; p++)
                                {
                                    DataRow ddp = fg.Rows[p];

                                    string fh1 = ddp["發票"].ToString();
                                    int gy = fh1.IndexOf("__");
                                    if (gy == -1)
                                    {
                                        sb.Append(ddp["發票"].ToString() + "/");
                                    }
                    

                    
                                }
                                if (sb.Length > 0)
                                {
                                    sb.Remove(sb.Length - 1, 1);
                                    drtemp5["發票"] = sb.ToString();
                                }
                            }
  
                        }
                        DuplicateKey = 單號;
                          dtCostDD.Rows.Add(drtemp5);
                   }


                   ACME.VoucherRPT frm = new ACME.VoucherRPT();
                   frm.dt = dtCostDD;
                   frm.ShowDialog();
                  
                  
                }
      
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
       System.Data.DataTable GetVoucher()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select T0.DOCENTRY 單號,T0.TRANSID ID,T2.PYMNTGROUP TERM,");
            sb.Append(" T0.CARDNAME 客戶名稱,T0.CARDCODE 客戶編號,T0.DOCDATE 過帳日期,@startdate as 起始日,@enddate  as 結束日,");
            sb.Append(" ISNULL(T3.ACCOUNT,62410001) 科目,ISNULL(T5.ACCTNAME,'進口費用') 名稱,ISNULL(T4.OCRNAME,'進出口管理課') 部門,isnull(CAST(T3.debit AS INT),0) 借方,isnull(CAST(T6.DEBIT AS INT),0) 營業稅");
            sb.Append(" ,u_pc_bsinv 發票,Convert(varchar(10),t0.u_pc_bsdat,111) 發票日期,");
            sb.Append(" case  when T0.CARDNAME like '%關務局%' then Convert(varchar(10),t0.u_pc_bsdat+10,111) ");
            sb.Append(" when T0.CARDNAME like '%港務局%' then Convert(varchar(10),t0.u_pc_bsdat+25,111) else ");
            sb.Append(" case when substring(Convert(varchar(10),t0.u_pc_bsdat,112),7,2) > 25 ");
            sb.Append(" THEN Convert(varchar(10),DateAdd(d,T2.extradays,DATEADD(day, -1, convert(datetime,(convert(char(7),dateadd(month,2,t0.u_pc_bsdat),111)+'/1')))),111) ");
            sb.Append(" ELSE Convert(varchar(10),DateAdd(d,T2.extradays,DATEADD(day, -1, convert(datetime,(convert(char(7),dateadd(month,1,t0.u_pc_bsdat),111)+'/1')))),111) end END 到期日");
            sb.Append(" from opch t0");
            sb.Append(" left join OCTG T2 ON (T0.GROUPNUM=T2.GROUPNUM)");
            sb.Append(" left join JDT1 T3 ON (T0.TRANSID=T3.TRANSID and  T3.debit <> 0 and T3.account <> '12640101') ");
            sb.Append(" left join OOCR T4 ON (T3.PROFITCODE=T4.OCRCODE)");
            sb.Append(" left join OACT T5 ON (T3.ACCOUNT=T5.ACCTCODE)");
            sb.Append(" left join (select TRANSID,sum(DEBIT) DEBIT from JDT1 where account = '12640101' group by TRANSID) T6 ON (T0.TRANSID=T6.TRANSID )");
            sb.Append(" where  1=1 ");
           if(textBox5.Text !="")
           {
            sb.Append("      and   case  when T0.CARDNAME like '%關務局%' then Convert(varchar(10),t0.u_pc_bsdat+10,112) ");
            sb.Append("            when T0.CARDNAME like '%港務局%' then Convert(varchar(10),t0.u_pc_bsdat+25,112) else ");
            sb.Append("            case when substring(Convert(varchar(10),t0.u_pc_bsdat,112),7,2) > 25 ");
            sb.Append("            THEN Convert(varchar(10),DateAdd(d,T2.extradays,DATEADD(day, -1, convert(datetime,(convert(char(7),dateadd(month,2,t0.u_pc_bsdat),111)+'/1')))),112) ");
            sb.Append("            ELSE Convert(varchar(10),DateAdd(d,T2.extradays,DATEADD(day, -1, convert(datetime,(convert(char(7),dateadd(month,1,t0.u_pc_bsdat),111)+'/1')))),112) end END between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "'");
           }
               if (textBox3.Text != "" && textBox4.Text != "")
            {

                sb.Append("  and T0.TRANSID between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "'  ");

            }
            else
            {
                sb.Append("  and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate  ");

            }
               if (textBox8.Text != "")
               {
                   sb.Append("  and ISNULL(T3.ACCOUNT,62410001)= '" + textBox8.Text.ToString() + "'  ");
               }
            sb.Append("   AND T0.CARDCODE=@tt  ");
            sb.Append("  union all  ");
            sb.Append(" select T0.DOCENTRY 單號,T0.TRANSID ID,T2.PYMNTGROUP TERM,");
            sb.Append(" T0.CARDNAME 客戶名稱,T0.CARDCODE 客戶編號,T0.DOCDATE 過帳日期,@startdate as 起始日,@enddate  as 結束日,");
            sb.Append(" ISNULL(T3.ACCOUNT,62410001) 科目,ISNULL(T5.ACCTNAME,'進口費用') 名稱,ISNULL(T4.OCRNAME,'進出口管理課') 部門,isnull(CAST(T3.credit AS INT),0)*-1 借方,isnull(CAST(T6.credit AS INT),0)*-1 營業稅");
            sb.Append(" ,u_rp_bsren 發票,Convert(varchar(10),t0.u_rp_bsdat,111) 發票日期,");
            sb.Append(" case  when T0.CARDNAME like '%關務局%' then Convert(varchar(10),t0.u_rp_bsdat+10,111) ");
            sb.Append(" when T0.CARDNAME like '%港務局%' then Convert(varchar(10),t0.u_rp_bsdat+25,111) else ");
            sb.Append(" case when substring(Convert(varchar(10),t0.u_rp_bsdat,112),7,2) > 25 ");
            sb.Append(" THEN Convert(varchar(10),DateAdd(d,T2.extradays,DATEADD(day, -1, convert(datetime,(convert(char(7),dateadd(month,2,t0.u_rp_bsdat),111)+'/1')))),111) ");
            sb.Append(" ELSE Convert(varchar(10),DateAdd(d,T2.extradays,DATEADD(day, -1, convert(datetime,(convert(char(7),dateadd(month,1,t0.u_rp_bsdat),111)+'/1')))),111) end END 到期日");
            sb.Append(" from orpc t0");
            sb.Append(" left join OCTG T2 ON (T0.GROUPNUM=T2.GROUPNUM)");
            sb.Append(" left join JDT1 T3 ON (T0.TRANSID=T3.TRANSID and  T3.credit <> 0 and T3.account <> '12640101') ");
            sb.Append(" left join OOCR T4 ON (T3.PROFITCODE=T4.OCRCODE)");
            sb.Append(" left join OACT T5 ON (T3.ACCOUNT=T5.ACCTCODE)");
            sb.Append(" left join (select TRANSID,sum(credit) credit from JDT1 where account = '12640101' group by TRANSID) T6 ON (T0.TRANSID=T6.TRANSID )");
            sb.Append(" where  1=1 ");

            if (textBox5.Text != "")
            {
                sb.Append("         and   case  when T0.CARDNAME like '%關務局%' then Convert(varchar(10),t0.u_rp_bsdat+10,112) ");
                sb.Append("            when T0.CARDNAME like '%港務局%' then Convert(varchar(10),t0.u_rp_bsdat+25,112) else ");
                sb.Append("            case when substring(Convert(varchar(10),t0.u_rp_bsdat,112),7,2) > 25 ");
                sb.Append("            THEN Convert(varchar(10),DateAdd(d,T2.extradays,DATEADD(day, -1, convert(datetime,(convert(char(7),dateadd(month,2,t0.u_rp_bsdat),111)+'/1')))),112) ");
                sb.Append("            ELSE Convert(varchar(10),DateAdd(d,T2.extradays,DATEADD(day, -1, convert(datetime,(convert(char(7),dateadd(month,1,t0.u_rp_bsdat),111)+'/1')))),112) end END between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "'");
            }

            if (textBox3.Text != "" && textBox4.Text != "")
            {

                sb.Append("  and T0.TRANSID between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "'  ");

            }
            else
            {
                sb.Append("  and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate  ");

            }

            if (textBox8.Text != "")
            {
                sb.Append("  and ISNULL(T3.ACCOUNT,62410001)= '" + textBox8.Text.ToString() + "'  ");
            }
            sb.Append("   AND T0.CARDCODE=@tt  ORDER BY T0.TRANSID "); 
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@tt", cardCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@startdate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@enddate", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        System.Data.DataTable GetVoucher2(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT U_PC_BSINV 發票,U_BSREN FROM [@CADMEN_FMD] T0 LEFT JOIN [@CADMEN_FMD1] T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append(" where  T0.U_BSREN = @DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();


            comboBox1.Text = "041-1";
        }



        private System.Data.DataTable MakeTableFcstWeek()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("TERM", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("科目", typeof(string));
            dt.Columns.Add("名稱", typeof(string));
            dt.Columns.Add("部門", typeof(string));
            dt.Columns.Add("借方", typeof(int));
            dt.Columns.Add("營業稅", typeof(int));
            dt.Columns.Add("起始日", typeof(string));
            dt.Columns.Add("結束日", typeof(string));
            dt.Columns.Add("發票", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("到期日", typeof(string));
            dt.Columns.Add("DUEDATE", typeof(string));
            dt.Columns.Add("單號", typeof(string));
            return dt;

        }

        private System.Data.DataTable MakeTableF()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            
        //    dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("號碼", typeof(int));
            dt.Columns.Add("傳票號碼", typeof(string));
            dt.Columns.Add("傳票日期", typeof(string));
            dt.Columns.Add("借貸", typeof(string));
            dt.Columns.Add("部門", typeof(string));
            dt.Columns.Add("科目編號", typeof(string));
            dt.Columns.Add("科目名稱", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("借方金額", typeof(decimal));
            dt.Columns.Add("貸方金額", typeof(decimal));
            dt.Columns.Add("公司", typeof(string));
            dt.Columns.Add("應付費用", typeof(decimal));
            dt.Columns.Add("手續費", typeof(decimal));
   
            return dt;

        }
        private void button2_Click(object sender, EventArgs e)
        {
            int n;
            if (!int.TryParse(textBox9.Text, out n) || !int.TryParse(textBox10.Text, out n))
            {
                MessageBox.Show("金額請輸入數字");
                return;
            }

            System.Data.DataTable GF = null;
            if (checkBox1.Checked)
            {
                GF = MakeTableF();
                string[] arrurl = textBox7.Text.Trim().Replace(System.Environment.NewLine, "").Split(new Char[] { ',' });
                int M1 = 0;
                foreach (string F in arrurl)
                {

                    M1++;
                    DataRow dr = null;

                    System.Data.DataTable GF2 = OJDT(F);
                    if (GF2.Rows.Count > 0)
                    {
                        for (int i = 0; i <= GF2.Rows.Count - 1; i++)
                        {

                            dr = GF.NewRow();

                            dr["號碼"] = M1;
                            dr["傳票號碼"] = GF2.Rows[i]["傳票號碼"].ToString();
                            dr["傳票日期"] = GF2.Rows[i]["傳票日期"].ToString();
                            dr["借貸"] = GF2.Rows[i]["借貸"].ToString();
                            dr["部門"] = GF2.Rows[i]["部門"].ToString();
                            dr["科目編號"] = GF2.Rows[i]["科目編號"].ToString();
                            dr["科目名稱"] = GF2.Rows[i]["科目名稱"].ToString();
                            dr["摘要"] = GF2.Rows[i]["摘要"].ToString();
                            dr["借方金額"] = Convert.ToDecimal(GF2.Rows[i]["借方金額"]);
                            dr["貸方金額"] = Convert.ToDecimal(GF2.Rows[i]["貸方金額"]);
                            dr["公司"] = GF2.Rows[i]["公司"].ToString();
                            dr["應付費用"] = Convert.ToDecimal(GF2.Rows[i]["應付費用"]);
                            dr["手續費"] = Convert.ToDecimal(GF2.Rows[i]["手續費"]);

                            GF.Rows.Add(dr);
                        }
                    }

                }


            }
            else
            {
                GF = OJDT("");
            }

           // dataGridView1.DataSource = GF;
            if (comboBox1.Text == "041-1")
            {

                if (textBox9.Text != "0" || textBox10.Text != "0")
                {
                    ACME.FormRafaF frm = new ACME.FormRafaF();
                    frm.dt = GF;
                    frm.ShowDialog();
                }
                else
                {
                    ACME.FormRafa frm = new ACME.FormRafa();
                    frm.dt = GF;
                    frm.ShowDialog();
                }

            }
            if (comboBox1.Text == "307")
            {
                if (textBox9.Text != "0" || textBox10.Text != "0")
                {
                    ACME.FormRafa2F frm = new ACME.FormRafa2F();
                    frm.dt = GF;
                    frm.ShowDialog();
                }
                else
                {
                    ACME.FormRafa2 frm = new ACME.FormRafa2();
                    frm.dt = GF;
                    frm.ShowDialog();
                }
            }
        }
        System.Data.DataTable OJDT(string transid)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT '傳票日期: '+convert(varchar,T0.refdate, 102) 傳票日期,cast(T0.TRANSID as varchar) 傳票號碼,cast(T0.TRANSID as varchar) 號碼,CASE ISNULL(DEBIT,0) WHEN 0 THEN '貸' ELSE '借' END 借貸, ");
            sb.Append("               PROFITCODE 部門,ACCOUNT 科目編號,T2.ACCTNAME 科目名稱,linememo 摘要, ");
            sb.Append("                DEBIT  借方金額,CREDIT 貸方金額,(SELECT COMPNYNAME  FROM OADM) 公司,@D1 應付費用,@D2 手續費   FROM ojdt T0 ");
            sb.Append("               LEFT JOIN JDT1 T1 ON(T0.TRANSID=T1.TRANSID) ");
            sb.Append("               LEFT JOIN OACT T2 ON(T1.ACCOUNT=T2.ACCTCODE) ");
            if (checkBox1.Checked)
            {
                sb.Append("  where T0.transid =@transid   ");
            }
            else
            {
                sb.Append("  where T0.transid in ( " + textBox7.Text + ") ORDER BY T0.transid  ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@D1", Convert.ToInt32(textBox9.Text)));
            command.Parameters.Add(new SqlParameter("@D2", Convert.ToInt32(textBox10.Text)));
            command.Parameters.Add(new SqlParameter("@transid", transid));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }




    }
}