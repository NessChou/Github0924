using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

namespace ACME
{
    public class ExcelUtils
    {
        /// <summary>   
        ///��k�A�ɥXDataGridView������ƨ�Excel��   
        /// </summary>   
        /// <remarks>  
        /// add com "Microsoft Excel 11.0 Object Library"  
        /// using Excel=Microsoft.Office.Interop.Excel;  
        /// using System.Reflection;  
        /// </remarks>  
        /// <param name= "dgv"> DataGridView </param>   
        /// 
        public static int DataGridViewToExcel(DataGridView dgv)
        {

            //int item=0;
            //�ө��O�s��ܤ��
            SaveFileDialog dlg = new SaveFileDialog();
            //�q�{�����X
            dlg.DefaultExt = "xlsx";
            //�����X�C��
            dlg.Filter = "EXCEL���(*.XLSX)|*.xlsx";
            //�q�M���|�O�t�η�e���|   
            dlg.InitialDirectory = Directory.GetCurrentDirectory();
            //���}�O�s��ܤ��
            if (dlg.ShowDialog() == DialogResult.Cancel) return 0;
            //��^�ɸ��|   
            string fileNameString = dlg.FileName;
            //����strFileName�O�_���ũέȵL��   
            if (fileNameString.Trim() == " ")
            { return 0; }
            //�w�q��椺��ƪ���ƩM�C��   
            int rowscount = dgv.Rows.Count;
            int colscount = dgv.Columns.Count;
            //��ƥ����j��0   
            if (rowscount <= 0)
            {
                MessageBox.Show("�S����ƥi�ѫO�s ", "���� ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            //�C�ƥ����j��0   
            if (colscount <= 0)
            {
                MessageBox.Show("�S����ƥi�ѫO�s ", "���� ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            //��Ƥ��i�H�j��65536   
            if (rowscount > 65536)
            {
                MessageBox.Show("��ưO���ƤӦh(�̦h����W�L65536��)�A����O�s ", "���� ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            //�C�Ƥ��i�H�j��255   
            if (colscount > 255)
            {
                MessageBox.Show("��ưO����ƤӦh�A����O�s ", "���� ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            //���ҥHfileNameString�R�W���ɬO�_�s�b�A�p�G�s�b�R����   
            FileInfo file = new FileInfo(fileNameString);
            if (file.Exists)
            {
                try
                {
                    file.Delete();
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message, "�R������ ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                }
            }

            Microsoft.Office.Interop.Excel.Application objExcel = null;
            Microsoft.Office.Interop.Excel.Workbook objWorkbook = null;
            Microsoft.Office.Interop.Excel.Worksheet objsheet = null;

            try
            {
                //�ө���H   
                objExcel = new Microsoft.Office.Interop.Excel.Application();
                objWorkbook = objExcel.Workbooks.Add(Missing.Value);
                objsheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkbook.ActiveSheet;
                //�]�mEXCEL���i��   
                objExcel.Visible = false;

                //�VExcel���g�J��檺���Y   
                int displayColumnsCount = 1;
                for (int i = 0; i <= dgv.ColumnCount - 1; i++)
                {
                    if (dgv.Columns[i].Visible == true)
                    {
                        objExcel.Cells[1, displayColumnsCount] = dgv.Columns[i].HeaderText.Trim();
                        displayColumnsCount++;
                    }
                }
                //�]�m�i�ױ�   

                //progressBar1.Refresh();
                //progressBar1.Visible = true;
                //progressBar1.Minimum = 1;
                //progressBar1.Maximum = dgv.RowCount;
                //progressBar1.Step = 1;
                //�VExcel���v��v�C�g�J��椤�����   
                for (int row = 0; row <= dgv.RowCount - 1; row++)
                {
                    //this.progressBar1.PerformStep();
                    //this.label2.Text = (this.progressBar1.Value / this.progressBar1.Maximum) + "%";

                    displayColumnsCount = 1;
                    for (int col = 0; col < colscount; col++)
                    {
                        if (dgv.Columns[col].Visible == true)
                        {
                            try
                            {
                                objExcel.Cells[row + 2, displayColumnsCount] = dgv.Rows[row].Cells[col].Value.ToString().Trim();
                                displayColumnsCount++;
                            }
                            catch (Exception)
                            {

                            }

                        }
                    }
                }

                //���öi�ױ�   
                //tempProgressBar.Visible   =   false;   
                //�O�s��   
                objWorkbook.SaveAs(fileNameString, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value);
                return 1;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "ĵ�i ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }
            finally
            {
                //����Excel����   
                if (objWorkbook != null) objWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                if (objExcel.Workbooks != null) objExcel.Workbooks.Close();
                if (objExcel != null) objExcel.Quit();

                objsheet = null;
                objWorkbook = null;
                objExcel = null;
            }
        }

        /// <summary>
        /// �N�ƾھɤJExcel
        /// </summary>
        /// <param name="view"></param>
        public static void ToExel(DataGridView view)
        {
            if (view.Rows.Count == 0)
            {
                MessageBox.Show("��椤�S����ơA����ɥX�Ū�", "��T����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //�إ�Excel����
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Application.Workbooks.Add(true);
            excel.Visible = true;
            //�ͦ����W��
            for (int i = 0; i < view.ColumnCount; i++)
            {

                excel.Cells[1, i + 1] = view.Columns[i].HeaderText;
            }
            //��R�ƾ�
            for (int row = 0; row <= view.RowCount - 1; row++)
            {
                for (int column = 0; column < view.ColumnCount; column++)
                {

                    if (view[column, row].ValueType == typeof(string))
                    {
                        excel.Cells[row + 2, column + 1] = "'" + view[column, row].Value.ToString();
                    }
                    else
                    {
                        excel.Cells[row + 2, column + 1] = view[column, row].Value.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// ��DataGridView������� �ɥX��EXCEL         /// </summary>
        /// <param name="dg"></param>
        public static void Input_Excel(DataGridView dg)
        {
            Microsoft.Office.Interop.Excel._Worksheet Sht;
            Microsoft.Office.Interop.Excel._Workbook Bo;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Bo = excel.Application.Workbooks.Add(true);
            // excel.Visible = true;//excel�O�_���
            Sht = (Microsoft.Office.Interop.Excel.Worksheet)Bo.Sheets[1];
            //�g�J��ƨ�EXCEL
            int Rowed = 0;
            if (dg.AllowUserToAddRows == true)
            {
                for (int i = 0; i < dg.Rows.Count - 1; i++)
                {
                    for (int y = 1; y <= dg.ColumnCount; y++)
                    {
                        excel.Cells[1, y] = dg.Columns[y - 1].HeaderText;
                    }
                    Rowed++;
                    if (Rowed < 65000)
                    {
                        for (int lie = 0; lie < dg.ColumnCount; lie++)
                        {
                            excel.Cells[Rowed + 1, lie + 1] = Convert.ToString(dg[lie, i].Value);
                        }
                    }
                    else
                    {
                        Sht = (Microsoft.Office.Interop.Excel.Worksheet)Bo.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        Rowed = 0;
                        i--;
                    }
                }
            }
            else
            {
                for (int i = 0; i < dg.Rows.Count - 1; i++)
                {
                    for (int y = 1; y <= dg.ColumnCount; y++)
                    {
                        excel.Cells[1, y] = dg.Columns[y - 1].HeaderText;
                    }
                    Rowed++;
                    if (Rowed < 65000)
                    {
                        for (int lie = 0; lie < dg.ColumnCount; lie++)
                        {
                            excel.Cells[Rowed + 1, lie + 1] = Convert.ToString(dg[lie, i].Value);
                        }
                    }
                    else
                    {
                        Sht = (Microsoft.Office.Interop.Excel.Worksheet)Bo.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        Rowed = 0;
                        i--;
                    }
                }

            }
            excel.Visible = true;
        }

    }

}
