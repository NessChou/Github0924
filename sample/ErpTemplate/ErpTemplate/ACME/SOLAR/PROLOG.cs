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
    public partial class PROLOG : Form
    {
        public string N1;
        public PROLOG()
        {
            InitializeComponent();
        }

        private void PROLOG_Load(object sender, EventArgs e)
        {

                string F1 = N1;

                System.Data.DataTable T1 = GetBOMLOG(F1);

                dataGridView1.DataSource = T1;

                if (globals.GroupID.ToString().Trim() == "SOLAR_1")
                {
                    for (int i = 0; i <= dataGridView1.Rows.Count - 2; i++)
                    {

                        DataGridViewRow row;

                        row = dataGridView1.Rows[i];
                        string a0 = row.Cells["DOCTYPE"].Value.ToString().Trim();
                        if (a0 == "PV" || a0 == "INV")
                        {
                            if (i == 0)
                            {
                                dataGridView1.CurrentCell = dataGridView1.Rows[1].Cells[0];
                            }
                            dataGridView1.Rows[i].Visible = false;
                        }
                    }

                }
        }

        private System.Data.DataTable GetBOMLOG(string shippingcode)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select DOCTYPE,OCOST 舊值,NCOST 新值,UPUSER 修改者,UDATE  修改時間,VER 版本,FATHER 母件編號,ITEMCODE 子件編號 from dbo.SOLAR_PROBOM3 where shippingcode=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
    }
}
