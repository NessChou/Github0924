using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ReflectionIT.Common.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
namespace ACME
{
    public partial class ShipChart2 : Form
    {
        public ShipChart2()
        {
            InitializeComponent();
        }

        private void radioButton1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox1.Items.Add("一月");
            listBox1.Items.Add("二月");
            listBox1.Items.Add("三月");
            listBox1.Items.Add("四月");
            listBox1.Items.Add("五月");
            listBox1.Items.Add("六月");
            listBox1.Items.Add("七月");
            listBox1.Items.Add("八月");
            listBox1.Items.Add("九月");
            listBox1.Items.Add("十月");
            listBox1.Items.Add("十一月");
            listBox1.Items.Add("十二月");
        }

        private void button5_Click(object sender, EventArgs e)
        {

            try
            {
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    al.Add(listBox2.Items[i].ToString());
                }
                StringBuilder sb = new StringBuilder();



                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }

                sb.Remove(sb.Length - 1, 1);

     
                   
                        APChart1 frm = new APChart1();
                        frm.a = sb.ToString();
                        frm.p = "2";
                        frm.co = comboBox1.SelectedValue.ToString();
                        frm.Show();
                    
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
     

       
        private void APChart_Load(object sender, EventArgs e)
        {
          
        }



        public static DataTable GetOP2(string aa,string bb)
        {
            String strCn = "Data Source=acmesrv13;Initial Catalog=acmesqlsp;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("              select count(*) 筆數,max(CHART.ChiMonth)+case substring(CREATENAME,0,2) when 'D' THEN '蕃茄' WHEN 'M' THEN '瑪姬' WHEN 'J' THEN '君穎'  END  月份 ");
            sb.Append("                          ,MAX(CHART.COLOR1) COLOR1,MAX(CHART.COLOR2) COLOR2,MAX(CHART.COLOR3) COLOR3,MAX(CHART.COLOR4) COLOR4   ");
            sb.Append("                          from shipping_main ");
            sb.Append("                              LEFT JOIN CHART ON (CHART.[MONTH]=substring(shipping_main.shippingcode,7,2))");
            sb.Append("                             where createname in ('maggietsai','dalychou','jeanwei')");
            sb.Append("                               and  substring(shipping_main.shippingcode,3,4)=convert(varchar(4),Getdate(),112)");
            sb.Append("                             AND DATEPART(week,hdate) BETWEEN @aa and @bb                             ");
            sb.Append("                          group by substring(shippingcode,7,2),createname");
            sb.Append("                                     order by substring(shippingcode,7,2),createname");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            
            command.Parameters.Add(new SqlParameter());
            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["shipping_main"];
        }

    }
}