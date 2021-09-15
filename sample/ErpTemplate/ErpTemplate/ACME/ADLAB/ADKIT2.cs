using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Excel;
namespace ACME
{
    public partial class ADKIT2 : Form
    {
        string strCnSP = "Data Source=acmesap;Initial Catalog=acmesqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public string PublicString;
        public ADKIT2()
        {
            InitializeComponent();
        }

        private void aD_OITM2BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aD_OITM2BindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.aD);

        }



        private void aD_OITM2DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "check2")
                {
                    for (int j = 0; j <= 1; j++)
                    {


                        System.Data.DataTable dt1 = aD.AD_OITM2;
                        int i = e.RowIndex;
                        DataRow drw = dt1.Rows[i];

                        string aa = drw["path"].ToString();


                        System.Diagnostics.Process.Start(aa);


                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ADKIT2_Load(object sender, EventArgs e)
        {
            try
            {
                if (!String.IsNullOrEmpty(PublicString))
                {


                    this.aD_OITM2TableAdapter.Fill(this.aD.AD_OITM2, PublicString);
                }
       
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                if (aD_OITM2DataGridView.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請選擇單據");
                    return;
                }


                this.Validate();
                this.aD_OITM2BindingSource.EndEdit();
                this.tableAdapterManager.UpdateAll(this.aD);



                string server = "//acmesrv01//SAP_Share//TTAdvance//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);

                if (result == DialogResult.OK)
                {
                    MessageBox.Show(Path.GetFileName(opdf.FileName));
                    string file = opdf.FileName;
                    bool FF1 = getrma.UploadFile(file, server, false);
                    if (FF1 == false)
                    {
                        return;
                    }


                    DataGridViewRow row;

                    row = aD_OITM2DataGridView.SelectedRows[0];
                    string a1 = row.Cells["ID"].Value.ToString();
                    string a2 = filename;

                    string a3 = @"\\acmesrv01\SAP_Share\TTAdvance\" + filename;


                    Updatepath(a2, a3, a1);


                    aD_OITM2TableAdapter.Fill(aD.AD_OITM2, PublicString);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void Updatepath(string filename, string path, string ID)
        {
            SqlConnection connection = new SqlConnection(strCnSP);

            StringBuilder sb = new StringBuilder();
            sb.Append(" update AD_OITM2 set filename=@filename,[path]=@path where  ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));
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
        private void aD_OITM2DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["ITEMCODE"].Value = PublicString;

            this.aD_OITM2BindingSource.EndEdit();
        }


    }
}
