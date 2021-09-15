using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace ACME
{
    public partial class GanttViewer : UserControl
    {

        #region Data Members

        int _todayIndex = -1;
        List<Task> _taskList = null;

        #endregion

        #region Expose Event for open task

        /// <summary>
        /// delegate for open task
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">open task event args</param>
        public delegate void OnTaskOpenDelegate(object sender, OpenTaskEventArgs e);

        private event OnTaskOpenDelegate _onTaskOpenEvent = null;

        /// <summary>
        /// register/unregister open task event
        /// </summary>
        public event OnTaskOpenDelegate OnTaskOpenEvent
        {
            add
            {
                _onTaskOpenEvent = (OnTaskOpenDelegate)Delegate.Combine(_onTaskOpenEvent, value);

            }
            remove
            {
                _onTaskOpenEvent = (OnTaskOpenDelegate)Delegate.Remove(_onTaskOpenEvent, value);
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public GanttViewer()
        {
            InitializeComponent();
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Method to Load Tasks
        /// </summary>
        /// <param name="tasks">task list</param>
        public void LoadTasks(List<Task> tasks)
        {

            gvGantt.ClearData();

            _taskList = tasks;

            // bind title grid
            gvTitle.AutoGenerateColumns = false;
            gvTitle.DataSource = tasks;


            // Bind gantt chart
            BindGanttChartGrid(tasks);
        }

        #endregion

        #region Private Text

        /// <summary>
        /// Bind gantt chart with task list 
        /// </summary>
        /// <param name="Tasks">ttask list</param>
        private void BindGanttChartGrid(List<Task> Tasks)
        {

            DataTable dt = new DataTable();

            DateTime? minDate = GetMiniDate(Tasks);
            DateTime? maxDate = GetMaxDate(Tasks);

            if (minDate.HasValue)
                minDate = minDate.Value.AddDays(-10);

            if (maxDate.HasValue)
                maxDate = maxDate.Value.AddDays(10);


            if (maxDate.HasValue && minDate.HasValue)
            {
                DateTime colDate = minDate.Value;


                while (colDate <= maxDate.Value)
                {
                    dt.Columns.Add(colDate.ToShortDateString());
                    colDate = colDate.AddDays(1);
                }

            }
            _todayIndex = -1;
            foreach (Task task in Tasks)
            {
                DataRow r = dt.NewRow();
                int i = 0;
                foreach (DataColumn column in dt.Columns)
                {
                    
                    DateTime columnDate = Convert.ToDateTime(DateTime.ParseExact(column.ColumnName, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo));
                    //DateTime columnDate = Convert.ToDateTime(column.ColumnName);
                    bool busy = false;
                    if (columnDate >= task.StartDate && columnDate <= task.EndDate)
                        busy = true;

                    //if (busy)
                    //{
                    //    r[i] = task.TaskID;
                    //}
                    //else
                    //{
                    //    r[i] = 0;
                    //}

                    //Terry
                    if (busy)
                    {
                        if (task.Status == TaskStatus.Completed.ToString())
                        {
                            r[i] = "Y_" + task.TaskID.ToString();
                        }
                        else
                        {
                            r[i] = "N_" + task.TaskID.ToString();
                        }
                    }
                    else
                    {
                        r[i] = "0";
                    }


                    if (_todayIndex < 0 && DateTime.Now.ToShortDateString() == columnDate.ToShortDateString())
                    {
                        _todayIndex = i;
                    }

                    i++;


                }

                dt.Rows.Add(r);
            }

            gvGantt.DataSource = dt;




        }

        /// <summary>
        /// get minimun date from task lsit
        /// </summary>
        /// <param name="tasks">task list</param>
        /// <returns>mini date</returns>
        private DateTime? GetMiniDate(List<Task> tasks)
        {
            //DateTime? date = null;

            DateTime? date = DateTime.Now;

            if (tasks.Count > 0)
            {
                //date = (from d in tasks select d.StartDate).Min();

                

                for (int i=0;i<=tasks.Count -1;i++)
                {
                  if (tasks[i].StartDate < date)
                  {
                    date  =tasks[i].StartDate;
                  }
                
                }


                //date = DateTime.Now.AddDays(-90);
                
            }

            return date;

        }

        /// <summary>
        /// get maximun date from task list
        /// </summary>
        /// <param name="tasks">task list</param>
        /// <returns>max date</returns>
        private DateTime? GetMaxDate(List<Task> tasks)
        {
          //  DateTime? date = null;

            DateTime? date = DateTime.Now;

            if (tasks.Count > 0)
            {
               // date = (from d in tasks select d.EndDate).Max();


                for (int i = 0; i <= tasks.Count - 1; i++)
                {
                    if (tasks[i].EndDate > date)
                    {
                        date = tasks[i].EndDate;
                    }

                }

               // date = DateTime.Now.AddDays(90);
            }

            return date;

        }

        /// <summary>
        /// get tool tip text
        /// </summary>
        /// <param name="taskID">task id</param>
        /// <returns>return tool tip text</returns>
        private string GetToolTipText(int taskID)
        {
            Task task = null;
            string toolTip = string.Empty;

            if (taskID > 0)
            {
                if (_taskList != null)
                {
                   // task = _taskList.Where(t => t.TaskID == taskID).First();
                }
            }


            if (task != null)
            {
                toolTip = "Start Date: " + task.StartDate.ToShortDateString() + "\n" +
                          "End Date: " + task.EndDate.ToShortDateString() + "\n" +
                          "Completed : " + task.PercentComplete.ToString() + "% \n" +
                          "Status : " + task.Status + "\n" +
                          "Priority : " + task.Priority + "\n";
            }

            return toolTip;
        }

        /// <summary>
        /// format priority cell format
        /// </summary>
        /// <param name="e">event args</param>
        private void FormatPriorityCell(DataGridViewCellFormattingEventArgs e)
        {
            // Ensure that the value is a string.
            String stringValue = e.Value as string;
            if (stringValue == null) return;


            // Set the cell ToolTip to the text value.
            DataGridViewCell cell = gvTitle[e.ColumnIndex, e.RowIndex];



            cell.ToolTipText = stringValue;


            // Replace the string value with the image value.
            switch (stringValue.ToLower())
            {
                case "high":
                    e.Value = "High";
                    break;

                case "low":
                    e.Value = "Low";
                    break;
                case "medium":
                    e.Value = "Medium";
                    break;

                default:
                    e.Value = "Low";
                    break;
            }

        }

        /// <summary>
        /// format color cell format
        /// </summary>
        /// <param name="e">event args</param>
        private void FormatColorCell(DataGridViewCellFormattingEventArgs e)
        {
            // Ensure that the value is a string.
            String stringValue = e.Value as string;
            if (stringValue == null) return;


            // Set the cell ToolTip to the text value.
            DataGridViewCell cell = gvTitle[e.ColumnIndex, e.RowIndex];



            cell.ToolTipText = stringValue;

            //if (!string.IsNullOrEmpty(stringValue))
            //{
            //    TaskCategory category = (TaskCategory)Enum.Parse(typeof(TaskCategory), stringValue);

            //    // Replace the string value with the image value.
            //    switch (category)
            //    {
            //        case TaskCategory.Blue:
            //            e.Value = Properties.Resources.BlueCat;
            //            break;

            //        case TaskCategory.Green:
            //            e.Value = Properties.Resources.GreenCat;
            //            break;

            //        case TaskCategory.Orange:
            //            e.Value = Properties.Resources.OrangeCat;
            //            break;
            //        case TaskCategory.Purple:
            //            e.Value = Properties.Resources.PurpleCat;
            //            break;
            //        case TaskCategory.Red:
            //            e.Value = Properties.Resources.RedCat;
            //            break;
            //        case TaskCategory.Yellow:
            //            e.Value = Properties.Resources.YellowCat;
            //            break;

            //        default:
            //            e.Value = Properties.Resources.Empty;
            //            break;
            //    }
            //}
            //else
            //{
            //    e.Value = Properties.Resources.Empty;
            //}

        }

        /// <summary>
        /// format flag cell format
        /// </summary>
        /// <param name="e">event args</param>
        private void FormatFlagCell(DataGridViewCellFormattingEventArgs e)
        {
            // Ensure that the value is a string.
            String stringValue = e.Value as string;
            if (stringValue == null) return;


            // Set the cell ToolTip to the text value.
            DataGridViewCell cell = gvTitle[e.ColumnIndex, e.RowIndex];



            cell.ToolTipText = stringValue;

            if (!string.IsNullOrEmpty(stringValue))
            {
                TaskFlag flag = (TaskFlag)Enum.Parse(typeof(TaskFlag), stringValue);

                // Replace the string value with the image value.
                e.Value = Utility.GetTaskFalgImage(flag);
            }
            else
            {
                //e.Value = Properties.Resources.Empty;
                e.Value = null;
            }

        }


        #endregion

        #region Events

        private void gvGantt_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (_todayIndex > 0)
            {
                gvGantt.FirstDisplayedScrollingColumnIndex = _todayIndex;
                gvGantt.FirstDisplayedScrollingRowIndex = 0;
            }

        }

        private void gvTitle_Scroll(object sender, ScrollEventArgs e)
        {
            gvGantt.FirstDisplayedScrollingRowIndex = gvTitle.FirstDisplayedScrollingRowIndex;

        }

        private void gvGantt_Scroll(object sender, ScrollEventArgs e)
        {
            gvTitle.FirstDisplayedScrollingRowIndex = gvGantt.FirstDisplayedScrollingRowIndex;
        }

        private void gvTitle_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {

            if (e.RowIndex < 0 && e.ColumnIndex >= 0)
            {
                Utility.DrawGradientRectangle(e.Graphics, e.CellBounds, Color.LightGray);
                e.PaintContent(e.CellBounds);
                e.Handled = true;
            }
        }

        private void gvGantt_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            int taskID = gvGantt.GetCellValue(e.RowIndex, e.ColumnIndex);

            e.ToolTipText = GetToolTipText(taskID);
        }

        private void gvTitle_DoubleClick(object sender, EventArgs e)
        {
            //if (gvTitle.SelectedCells.Count > 0)
            //{
            //    int rowIndex = gvTitle.SelectedCells[0].RowIndex;
            //    Task task = gvTitle.Rows[rowIndex].DataBoundItem as Task;

            //    if (task != null && _onTaskOpenEvent != null)
            //    {
            //        _onTaskOpenEvent(gvTitle, new OpenTaskEventArgs(task.TaskID));
            //    }

            //}

        }

        private void gvTitle_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewColumn column = gvTitle.Columns[e.ColumnIndex];

            if (column.Name.Equals("TaskPriority"))
            {
                FormatPriorityCell(e);
            }
            else if (column.Name.Equals("TaskColorCategory"))
            {
                FormatColorCell(e);
            }
            else if (column.Name.Equals("TaskFlag"))
            {
                FormatFlagCell(e);
            }

        }

        #endregion

    }
}


