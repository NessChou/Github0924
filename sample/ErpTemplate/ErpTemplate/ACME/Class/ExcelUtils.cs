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
        ///方法，導出DataGridView中的資料到Excel檔   
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
            //申明保存對話方塊
            SaveFileDialog dlg = new SaveFileDialog();
            //默認文件尾碼
            dlg.DefaultExt = "xlsx";
            //文件尾碼列表
            dlg.Filter = "EXCEL文件(*.XLSX)|*.xlsx";
            //默然路徑是系統當前路徑   
            dlg.InitialDirectory = Directory.GetCurrentDirectory();
            //打開保存對話方塊
            if (dlg.ShowDialog() == DialogResult.Cancel) return 0;
            //返回檔路徑   
            string fileNameString = dlg.FileName;
            //驗證strFileName是否為空或值無效   
            if (fileNameString.Trim() == " ")
            { return 0; }
            //定義表格內資料的行數和列數   
            int rowscount = dgv.Rows.Count;
            int colscount = dgv.Columns.Count;
            //行數必須大於0   
            if (rowscount <= 0)
            {
                MessageBox.Show("沒有資料可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            //列數必須大於0   
            if (colscount <= 0)
            {
                MessageBox.Show("沒有資料可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            //行數不可以大於65536   
            if (rowscount > 65536)
            {
                MessageBox.Show("資料記錄數太多(最多不能超過65536條)，不能保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            //列數不可以大於255   
            if (colscount > 255)
            {
                MessageBox.Show("資料記錄行數太多，不能保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            //驗證以fileNameString命名的檔是否存在，如果存在刪除它   
            FileInfo file = new FileInfo(fileNameString);
            if (file.Exists)
            {
                try
                {
                    file.Delete();
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message, "刪除失敗 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                }
            }

            Microsoft.Office.Interop.Excel.Application objExcel = null;
            Microsoft.Office.Interop.Excel.Workbook objWorkbook = null;
            Microsoft.Office.Interop.Excel.Worksheet objsheet = null;

            try
            {
                //申明對象   
                objExcel = new Microsoft.Office.Interop.Excel.Application();
                objWorkbook = objExcel.Workbooks.Add(Missing.Value);
                objsheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkbook.ActiveSheet;
                //設置EXCEL不可見   
                objExcel.Visible = false;

                //向Excel中寫入表格的表頭   
                int displayColumnsCount = 1;
                for (int i = 0; i <= dgv.ColumnCount - 1; i++)
                {
                    if (dgv.Columns[i].Visible == true)
                    {
                        objExcel.Cells[1, displayColumnsCount] = dgv.Columns[i].HeaderText.Trim();
                        displayColumnsCount++;
                    }
                }
                //設置進度條   

                //progressBar1.Refresh();
                //progressBar1.Visible = true;
                //progressBar1.Minimum = 1;
                //progressBar1.Maximum = dgv.RowCount;
                //progressBar1.Step = 1;
                //向Excel中逐行逐列寫入表格中的資料   
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

                //隱藏進度條   
                //tempProgressBar.Visible   =   false;   
                //保存檔   
                objWorkbook.SaveAs(fileNameString, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value);
                return 1;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }
            finally
            {
                //關閉Excel應用   
                if (objWorkbook != null) objWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                if (objExcel.Workbooks != null) objExcel.Workbooks.Close();
                if (objExcel != null) objExcel.Quit();

                objsheet = null;
                objWorkbook = null;
                objExcel = null;
            }
        }

        /// <summary>
        /// 將數據導入Excel
        /// </summary>
        /// <param name="view"></param>
        public static void ToExel(DataGridView view)
        {
            if (view.Rows.Count == 0)
            {
                MessageBox.Show("表格中沒有資料，不能導出空表", "資訊提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //建立Excel物件
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Application.Workbooks.Add(true);
            excel.Visible = true;
            //生成欄位名稱
            for (int i = 0; i < view.ColumnCount; i++)
            {

                excel.Cells[1, i + 1] = view.Columns[i].HeaderText;
            }
            //填充數據
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
        /// 把DataGridView中的資料 導出到EXCEL         /// </summary>
        /// <param name="dg"></param>
        public static void Input_Excel(DataGridView dg)
        {
            Microsoft.Office.Interop.Excel._Worksheet Sht;
            Microsoft.Office.Interop.Excel._Workbook Bo;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Bo = excel.Application.Workbooks.Add(true);
            // excel.Visible = true;//excel是否顯示
            Sht = (Microsoft.Office.Interop.Excel.Worksheet)Bo.Sheets[1];
            //寫入資料到EXCEL
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
