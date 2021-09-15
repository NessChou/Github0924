using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Drawing.Drawing2D;
using System.Globalization;


namespace ACME
{
    public partial class GanttChart : System.Windows.Forms.DataGridView
    {

        #region Data Members

        Color _headerColor = Color.LightGray;
        Color _weekendColor = Color.LightGray;
        Color _todayColor = Color.Green;
        Color _taskColor = Color.Red;

        Color _EndColor = Color.Green;       

        int Default_Cell_Width = 25;

        #endregion

        #region Internal Class

        /// <summary>
        /// Class to get the Header Cell data
        /// </summary>
        internal class HeaderCellData
        {
            #region Data Members

            int _month = 0;
            int _days = 0;
            int _year = 0;
            string _title = string.Empty;
            string _columnIndex = string.Empty;

            public string ColumnIndexes
            {
                get
                {

                    return _columnIndex;
                }
                set { _columnIndex = value; }
            }

            #endregion

            #region Properties

            public int Month
            {
                get { return _month; }
                set { _month = value; }
            }

            public int Days
            {
                get { return _days; }
                set { _days = value; }
            }

            public int Year
            {
                get { return _year; }
                set { _year = value; }
            }

            public string Title
            {
                get { return _title; }
                set { _title = value; }
            }

            #endregion


            public HeaderCellData(int month, int days, int year, string title, string columnIndexes)
            {
                _month = month;
                _days = days;
                _year = year;
                _title = title;
                _columnIndex = columnIndexes;
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public GanttChart()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// Data source data table
        /// </summary>
        private DataTable DataSourceTable
        {
            get
            {
                return DataSource as DataTable;
            }
        }

        /// <summary>
        /// get/set task color
        /// </summary>
        public Color TaskColor
        {
            get { return _taskColor; }
            set { _taskColor = value; }
        }

        /// <summary>
        /// get/sset today color
        /// </summary>
        public Color TodayColor
        {
            get { return _todayColor; }
            set { _todayColor = value; }
        }

        /// <summary>
        /// get/set weekend color
        /// </summary>
        public Color WeekendColor
        {
            get { return _weekendColor; }
            set { _weekendColor = value; }
        }

        /// <summary>
        /// get/set header color
        /// </summary>
        public Color HeaderColor
        {
            get { return _headerColor; }
            set { _headerColor = value; }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Method to clear calendar data
        /// </summary>
        public void ClearData()
        {
            for (int i = Columns.Count - 1; i >= 0; i--)
            {
                Columns.Remove(Columns[i]);
            }

            this.Invalidate();
        }

        /// <summary>
        /// Method to get cell value
        /// </summary>
        /// <param name="rowIndex">index of row</param>
        /// <param name="columnIndex">index of column</param>
        /// <returns>cell value</returns>
        public int GetCellValue(int rowIndex, int columnIndex)
        {
            int id = 0;
            if (rowIndex >= 0 && columnIndex >= 0)
            {
                try
                {
                    id = Convert.ToInt32(DataSourceTable.Rows[rowIndex][columnIndex]);
                }
                catch { }
            }
            return id;
        }

        #endregion

        #region Overrided Events

        protected override void OnCellPainting(DataGridViewCellPaintingEventArgs e)
        {
            // Draw row selector header
            if (e.ColumnIndex < 0)
            {
                DrawCellColor(e.Graphics, e.CellBounds, Color.SteelBlue, Color.Black);
                e.PaintContent(e.CellBounds);
            }
            /// Draw Header Cell
            else if (e.RowIndex < 0 && e.ColumnIndex >= 0)
            {
                DrawDayHeaderCell(e); 
            }
            /// Draw Cell
            else
            {
                DrawDataCell(e);
                
            }

            // set true to set it is handled
            e.Handled = true;

        }
            
        protected override void OnPaint(PaintEventArgs e)
        {

            DrawYearTopHeader(e.Graphics);

            base.OnPaint(e);
            

        }

        protected override void OnColumnAdded(DataGridViewColumnEventArgs e)
        {
            e.Column.Width = Default_Cell_Width;
            base.OnColumnAdded(e);
        }

        protected override void OnColumnWidthChanged(DataGridViewColumnEventArgs e)
        {
            // on column width change. change the width of all columns
            int width = e.Column.Width;

            foreach (DataGridViewColumn col in Columns)
            {
                col.Width = width;
            }

            this.Invalidate();
            base.OnColumnWidthChanged(e);

        }

        protected override void OnRowHeightChanged(DataGridViewRowEventArgs e)
        {

            // On row height change, change the height of all rows
            int height = e.Row.Height;

            foreach (DataGridViewRow row in Rows)
            {
                row.Height = height;
            }

            this.Invalidate();

            base.OnRowHeightChanged(e);
        }

        protected override void OnScroll(ScrollEventArgs e)
        {

            Rectangle r = new Rectangle();

            r.X = 0;
            r.Y = 0;
            r.Width = this.Width;
            r.Height = this.ColumnHeadersHeight;

            Invalidate(r);

            base.OnScroll(e);

        }
        
        #endregion

        #region Private Methods
        
        /// <summary>
        /// Get months name list from data source
        /// </summary>
        /// <returns>list of months</returns>
        private List<HeaderCellData> GetMonthName()
        {
            List<HeaderCellData> names = new List<HeaderCellData>();

            DataTable dtMain = new DataTable();
            dtMain = DataSourceTable;

            if (dtMain != null)
            {
                int days = 0;


                DateTime lastDate = DateTime.MinValue; ;
                string index = "";
                int i = 0;
                foreach (DataColumn column in dtMain.Columns)
                {
                    DateTime date = Convert.ToDateTime(DateTime.ParseExact(column.ColumnName, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo));
                
                    //  DateTime date = Convert.ToDateTime(column.ColumnName);

                    if (lastDate != DateTime.MinValue && lastDate.Month != date.Month)
                    {
                        names.Add(new HeaderCellData(lastDate.Month, days, lastDate.Year, lastDate.ToString("MMM, yyyy"), index));
                        days = 0;
                        index = string.Empty;
                    }

                    if (index.Length > 0)
                    {
                        index += ",";
                    }

                    index += i.ToString();
                    days++;
                    lastDate = date;
                    i++;
                }

                if (days != 0)
                {
                    names.Add(new HeaderCellData(lastDate.Month, days, lastDate.Year, lastDate.ToString("MMM, yyyy"), index));

                }
            }

            return names;
        }

        /// <summary>
        /// Draw month month rectanlg of grid
        /// </summary>
        /// <param name="g">graphics</param>
        private void DrawYearTopHeader(Graphics g)
        {
            List<HeaderCellData> months = GetMonthName();

            for (int j = 0; j < months.Count; j++)
            {
                HeaderCellData monthDetail = months[j];


                string[] indexes = monthDetail.ColumnIndexes.Split(',');


                Rectangle rectangle1 = Rectangle.Empty;
                foreach (string index in indexes)
                {
                    Rectangle rect = GetCellRectangle(Convert.ToInt32(index), -1);
                    if (rect != Rectangle.Empty)
                    {
                        if (rectangle1 == Rectangle.Empty)
                        {
                            rectangle1 = rect;
                        }
                        else
                        {
                            rectangle1 = Rectangle.Union(rectangle1, rect);
                        }
                    }
                   
                }

                if (rectangle1 != Rectangle.Empty)
                {
                    rectangle1.Height = rectangle1.Height / 2;
                    DrawCellColor(g, rectangle1, _headerColor, Color.Black, true);
                    DrawString(g, monthDetail.Title, rectangle1);
                }


            }

        }
        
        /// <summary>
        /// Draw day header cell
        /// </summary>
        /// <param name="e">event Args</param>
        private void DrawDayHeaderCell(DataGridViewCellPaintingEventArgs e)
        {

            Rectangle r = e.CellBounds;
            r.Height = r.Height / 2;
            r.Y += e.CellBounds.Height / 2;

            DrawCellColor(e.Graphics, r, _headerColor, Color.Black, true);

            DataTable dt = DataSourceTable;
            DateTime colDate = GetColumnDate(e.ColumnIndex);
            string text = colDate.Day.ToString();

            if (IsTodayCell(colDate))
            {
                DrawTodayLine(e.Graphics, e.CellBounds);
            }

            DrawString(e.Graphics, text, r);

        }

        /// <summary>
        /// Draw data cell
        /// </summary>
        /// <param name="e">event args</param>
        private void DrawDataCell(DataGridViewCellPaintingEventArgs e)
        {
            DateTime colDate = GetColumnDate(e.ColumnIndex);
            Color backColor = (colDate.DayOfWeek == DayOfWeek.Saturday || colDate.DayOfWeek == DayOfWeek.Sunday) ? _weekendColor : Color.White;

            DrawCellColor(e.Graphics, e.CellBounds, backColor, Color.Black);

            if (IsTodayCell(colDate))
            {
                DrawTodayLine(e.Graphics, e.CellBounds);
            }

            // Show task color bar
            if (IsBusyCell(e.RowIndex, e.ColumnIndex))
            {
                Rectangle rect = new Rectangle(e.CellBounds.Location, e.CellBounds.Size);

                rect.Y += 2;
                rect.Height -= 5;

                //  DrawCellColor(e.Graphics, rect, Color.Red, Color.Red);

                rect.X -= 1;
                rect.Width += 1;
                Utility.DrawGradientRectangle(e.Graphics, rect, _taskColor);

                //r[i] = task.TaskID; //可以取得狀態
                //terrylee
                if (DataSourceTable.Rows[e.RowIndex][e.ColumnIndex].ToString().Substring(0, 1) == "Y")
                {
                    Utility.DrawGradientRectangle(e.Graphics, rect, _EndColor);
                }
            }

        }

        /// <summary>
        /// Draw cell color
        /// </summary>
        /// <param name="g">graphics</param>
        /// <param name="bounds">cell bounds</param>
        /// <param name="backColor">back Color</param>
        /// <param name="borderColor">border color</param>
        private void DrawCellColor(Graphics g, Rectangle bounds, Color backColor, Color borderColor)
        {
            DrawCellColor(g, bounds, backColor, borderColor, false);
        }
        
        /// <summary>
        /// Draw cell
        /// </summary>
        /// <param name="g">graphics</param>
        /// <param name="bounds">cell bounds</param>
        /// <param name="backColor">back color</param>
        /// <param name="borderColor">border coor</param>
        /// <param name="drawGradeient">true to fill gradient</param>
        private void DrawCellColor(Graphics g, Rectangle bounds, Color backColor, Color borderColor, bool drawGradeient)
        {
        
            using (Brush brush = new SolidBrush(backColor))
            {
                using (Pen p = new Pen(borderColor))
                {
                    

                    Rectangle rect1 = new Rectangle(bounds.Location, bounds.Size);
                    Rectangle rect2 = new Rectangle(bounds.Location, bounds.Size);

                    rect1.X -= 1;
                    rect1.Y -= 1;

                    rect2.Width -= 1;
                    rect2.Height -= 1;

                    // must draw border for grid scrolling horizontally 
                    g.DrawRectangle(p, rect1);

                    if (!drawGradeient)

                        g.FillRectangle(brush, rect2);
                    else
                        Utility.DrawGradientRectangle(g, rect2, backColor);
                }
            }


        }

        /// <summary>
        /// Draw string
        /// </summary>
        /// <param name="g">graphics</param>
        /// <param name="text">text to display</param>
        /// <param name="rectangle">bounds</param>
        private void DrawString(Graphics g, string text, Rectangle rectangle)
        {
            StringFormat format = new StringFormat();

            format.Alignment = StringAlignment.Center;

            format.LineAlignment = StringAlignment.Center;

            g.DrawString(text,

                this.ColumnHeadersDefaultCellStyle.Font,

                new SolidBrush(this.ColumnHeadersDefaultCellStyle.ForeColor),

                rectangle, format);

        }

        /// <summary>
        /// get bounds of cell
        /// </summary>
        /// <param name="columnIndex">column index</param>
        /// <param name="rowIndex">row index</param>
        /// <returns>bounds rectangle</returns>
        private Rectangle GetCellRectangle(int columnIndex, int rowIndex)
        {
            Rectangle rect = Rectangle.Empty;
            try
            {
                rect = this.GetCellDisplayRectangle(columnIndex, rowIndex, true);
            }
            catch
            {
            }

            return rect;
        }
        
        /// <summary>
        /// get column date
        /// </summary>
        /// <param name="columnIndex">column index</param>
        /// <returns>date time</returns>
        private DateTime GetColumnDate(int columnIndex)
        {
            
            return Convert.ToDateTime(Convert.ToDateTime(DateTime.ParseExact(DataSourceTable.Columns[columnIndex].ColumnName, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo)));
            // return Convert.ToDateTime(DataSourceTable.Columns[columnIndex].ColumnName);
        }

        /// <summary>
        /// get true is cell has task id
        /// </summary>
        /// <param name="rowIndex">row index</param>
        /// <param name="columnIndex">column index</param>
        /// <returns>true if cell has value</returns>
        private bool IsBusyCell(int rowIndex, int columnIndex)
        {
            bool busy = false;
            try
            {
                busy = (DataSourceTable.Rows[rowIndex][columnIndex].ToString() != "0");
            }
            catch { }
            return busy;
        }
        
        /// <summary>
        /// retrun true if datetime is today
        /// </summary>
        /// <param name="datetime">datetime</param>
        /// <returns>true for today date</returns>
        private bool IsTodayCell(DateTime datetime)
        {
            bool today = false;

            if (datetime.ToShortDateString() == DateTime.Now.ToShortDateString())
            {
                today = true;
            }

            return today;
        }

        /// <summary>
        /// draw today line in cell
        /// </summary>
        /// <param name="graphics">graphics</param>
        /// <param name="cellBounds">cell bounds</param>
        private void DrawTodayLine(Graphics graphics, Rectangle cellBounds)
        {

            Rectangle rect = new Rectangle();
            rect.X = (cellBounds.X + (cellBounds.Width / 2)) - 1;
            rect.Y = cellBounds.Y;
            rect.Width = 2;
            rect.Height = cellBounds.Height;

            DrawCellColor(graphics, rect, _todayColor, _todayColor);


        }

        #endregion

    }

}
