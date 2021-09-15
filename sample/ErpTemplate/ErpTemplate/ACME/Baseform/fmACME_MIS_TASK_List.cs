 using System;
 using System.Collections.Generic;
 using System.ComponentModel;
 using System.Data;
 using System.Drawing;
 using System.Text;
 using System.Windows.Forms;
 using System.Data.SqlClient;
 
 //using Microsoft.Office.Interop.Word;
 using System.IO;
using System.Reflection;

//http://support.microsoft.com/kb/316384

//http://www.c-sharpcorner.com/UploadFile/amrish_deep/WordAutomation05102007223934PM/WordAutomation.aspx

//find & Replace
//http://www.codeproject.com/KB/office/Word_Automation.aspx

namespace ACME
{
    public partial class fmACME_MIS_TASK_List : Form
    {

        public fmACME_MIS_TASK_List()
        {
            InitializeComponent();
            gvData.AutoGenerateColumns = false;
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt = GetACME_MIS_TASK_Condition();

            gvData.DataSource = dt;

        }

        private void gvData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int scrollPosition = e.RowIndex;

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewColumn column = (sender as DataGridView).Columns[e.ColumnIndex];
                if (column.Name == "colEdit")
                {

                    DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;
                    if (row != null)
                    {
                       
                        fmACME_MIS_TASK form = new fmACME_MIS_TASK(Convert.ToInt32(row["ID"]));
                        if (form.ShowDialog() == DialogResult.OK)
                        {
                            RefreshData();
                            try
                            {
                                (sender as DataGridView).CurrentCell = (sender as DataGridView)[0, scrollPosition];
                            }
                            catch
                            {

                            }
                        }

                    }
                }

            }
        }

        private void RefreshData()
        {
            System.Data.DataTable dt = GetACME_MIS_TASK_Condition();

            gvData.DataSource = dt;

        }

        private void btnAdd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            fmACME_MIS_TASK form = new fmACME_MIS_TASK(0);

            if (form.ShowDialog() == DialogResult.OK)
            {
                RefreshData();
            }
        }


        // Condition 版本
        public System.Data.DataTable GetACME_MIS_TASK_Condition()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;


            sb.Append("SELECT * FROM ACME_MIS_TASK WHERE  Owner ='lleytonchen' ");


            if (!string.IsNullOrEmpty(TextBoxStartDate1.Text))
            {
                sb.Append(" AND StartDate >=@StartDate1 ");
                command.Parameters.Add(new SqlParameter("@StartDate1", TextBoxStartDate1.Text));
            }

            if (!string.IsNullOrEmpty(TextBoxStartDate2.Text))
            {
                sb.Append(" AND StartDate <=@StartDate2 ");
                command.Parameters.Add(new SqlParameter("@StartDate2", TextBoxStartDate2.Text));
            }


            //未結
            if (radioButton1.Checked)
            {
                sb.Append(" and (AcDate is null  or AcDate='')");
            }

            //已結
            if (radioButton2.Checked)
            {
                sb.Append(" and (AcDate is not null  and  AcDate <>'') ");
            }


            if (!string.IsNullOrEmpty(TextBoxAcDate1.Text))
            {
                sb.Append(" AND AcDate >=@AcDate1 ");
                command.Parameters.Add(new SqlParameter("@AcDate1", TextBoxAcDate1.Text));
            }

            if (!string.IsNullOrEmpty(TextBoxAcDate2.Text))
            {
                sb.Append(" AND AcDate <=@AcDate2 ");
                command.Parameters.Add(new SqlParameter("@AcDate2", TextBoxAcDate2.Text));
            }
            sb.Append(" order by AcDate desc ");
            command.CommandText = sb.ToString();
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MIS_TASK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MIS_TASK"];
        }


        // Condition 版本
        public System.Data.DataTable GetACME_MIS_TASK_Progress()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;


            sb.Append("SELECT * FROM ACME_MIS_TASK WHERE 1= 1");


            sb.Append(" AND Owner ='lleytonchen' ");

                sb.Append(" and (AcDate is null  or AcDate='')");

                sb.Append(" order by EndDate");
           

          


            command.CommandText = sb.ToString();
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MIS_TASK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MIS_TASK"];
        }



        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }


        private string FormatDateStr(string sDate)
        {

            try
            {
                return sDate.Substring(0, 4) + "/" + sDate.Substring(4, 2) + "/" + sDate.Substring(6, 2);
            }
            catch
            {
                return "";
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
      
                TextBoxAcDate1.Text = DateTime.Now.AddDays(-7).ToString("yyyyMMdd");
                TextBoxAcDate2.Text = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");

                TextBoxStartDate1.Text = "";
                TextBoxStartDate2.Text = "";
                radioButton2.Checked = true;
            }
            else
            {

                TextBoxAcDate1.Text = "";
                TextBoxAcDate2.Text ="";

                TextBoxStartDate1.Text = "";
                TextBoxStartDate2.Text = "";
                radioButton3.Checked = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(gvData);
        }



        //private void CreateWord()
        //{
        //    Word.Application wdApp= new Word.Application();
        //    Word.Document wdDoc;
        //    Word.Table wdTab;
        //    DataTable dt;
        //    int rowIndex,colIndex;
        //    object oVisible = true;
        //    object oMissing = System.Reflection.Missing.Value;
        //    object oStart = 0;
        //    object oEnd = 0;
        //    wdDoc = wdApp.Documents.Add(ref oMissing,ref oMissing,ref oMissing,ref oVisible);
        //    rowIndex = 1;
        //    colIndex = 0;
        //    dt  = CreateTable();
        //    wdTab = wdDoc.Tables.Add(wdDoc.Range(ref oStart,ref oEnd),dt.Rows.Count +1,dt.Columns.Count,ref oMissing,ref oMissing);
        //    foreach(DataColumn Col in dt.Columns)
        //    {
        //        colIndex++;
        //        wdTab.Cell(1,colIndex).Range.InsertAfter(Col.ColumnName);
        //    }
        //    foreach(DataRow Row in dt.Rows)
        //    {
        //        rowIndex++;
        //        colIndex = 0;
        //        foreach(DataColumn Col in dt.Columns)
        //        {
        //            colIndex++;
        //            wdTab.Cell(rowIndex,colIndex).Range.InsertAfter(Row[Col.ColumnName].ToString());
        //        }
        //    }
        //    wdTab.Borders.InsideLineStyle  = Word.WdLineStyle.wdLineStyleSingle;
        //    wdTab.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
        //    wdApp.Visible = true;
        //}

        // Save opened template as another .doc file
//Object oSaveAsFile = Server.MapPath("") + "\\MyTemplate\\MyDocument.doc";
//oDoc.SaveAs(ref oSaveAsFile, ref oMissing, ref oMissing, ref oMissing,
//                   ref oMissing, ref oMissing, ref oMissing, ref oMissing, 
//                   ref oMissing, ref oMissing, ref oMissing);

//// Close the document, destroy the object and Quit Word
//object SaveChanges = true;
//oDoc.Close(ref SaveChanges, ref oMissing, ref oMissing);
//oDoc = null;
//oWord.Quit(ref SaveChanges, ref oMissing, ref oMissing);

    }
}

