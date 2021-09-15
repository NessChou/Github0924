using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace ACME
{

    /// <summary>
    /// class for general functionality
    /// </summary>
    public static class Utility
    {

        #region Properties

        /// <summary>
        /// get jump list display style.
        /// </summary>
        public static JumpListDisplayStyle JumpListDisplayStyle
        {
            get
            {
                try
                {
                    string strValue = "ShowTaskListBy";
                    return (JumpListDisplayStyle)Enum.Parse(typeof(JumpListDisplayStyle), strValue);
                }
                catch
                {

                    return JumpListDisplayStyle.ColorCategory;
                }
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Method to bind priority combobox
        /// </summary>
        /// <param name="comboBox">drop down</param>
        /// <param name="addEmpty">true for adding empty value add top</param>
        public static void BindTaskPriorityCombo(ComboBox priorityCombo, bool addEmpty)
        {
            priorityCombo.DrawMode = DrawMode.OwnerDrawVariable;
            priorityCombo.DrawItem += new DrawItemEventHandler(priorityCombo_DrawItem);
            priorityCombo.Items.Clear();

            if (addEmpty)
            {
                priorityCombo.Items.Add("");
            }


            foreach (TaskPriority priority in Enum.GetValues(typeof(TaskPriority)))
            {
                priorityCombo.Items.Add(priority);
            }

        }

        /// <summary>
        /// Method to bind color caegory drop down
        /// </summary>
        /// <param name="colorCategoryCombo">combo box</param>
        /// <param name="addEmpty">true to add empty item</param>
        public static void BindTaskColorCategoryCombo(ComboBox colorCategoryCombo, bool addEmpty)
        {
            colorCategoryCombo.DrawMode = DrawMode.OwnerDrawVariable;
            colorCategoryCombo.DrawItem += new DrawItemEventHandler(colorCategoryCombo_DrawItem);
            colorCategoryCombo.Items.Clear();

            if (addEmpty)
            {
                colorCategoryCombo.Items.Add("");
            }


            foreach (TaskCategory category in Enum.GetValues(typeof(TaskCategory)))
            {
                colorCategoryCombo.Items.Add(category);
            }

        }

        /// <summary>
        /// Method to bind task flag drop down
        /// </summary>
        /// <param name="comboBox">combo box</param>
        /// <param name="addEmpty">true to add empty item</param>        
        public static void BindTaskFlagCombo(ComboBox taskFlagCombo, bool addEmpty)
        {
            taskFlagCombo.DrawMode = DrawMode.OwnerDrawVariable;
            taskFlagCombo.DrawItem += new DrawItemEventHandler(taskFlagCombo_DrawItem);

            taskFlagCombo.Items.Clear();

            if (addEmpty)
            {
                taskFlagCombo.Items.Add("");
            }


            foreach (TaskFlag flag in Enum.GetValues(typeof(TaskFlag)))
            {
                taskFlagCombo.Items.Add(flag);
            }

        }

        /// <summary>
        /// Method to bind status drop down
        /// </summary>
        /// <param name="comboBox">drop down</param>
        /// <param name="addEmpty">true for adding empty value add top</param>
        public static void BindTaskStatusCombo(ComboBox comboBox, bool addEmpty)
        {

            comboBox.Items.Clear();

            if (addEmpty)
            {
                comboBox.Items.Add("");
            }


            foreach (TaskStatus status in Enum.GetValues(typeof(TaskStatus)))
            {
                comboBox.Items.Add(status);
            }

        }

        /// <summary>
        /// Method to bind task duration drop down
        /// </summary>
        /// <param name="comboBox">combo box</param>
        /// <param name="addEmpty">true to add empty item</param>
        public static void BindTaskDurationTypeCombo(ComboBox comboBox, bool addEmpty)
        {

            comboBox.Items.Clear();

            if (addEmpty)
            {
                comboBox.Items.Add("");
            }


            foreach (TaskDurationType durationType in Enum.GetValues(typeof(TaskDurationType)))
            {
                comboBox.Items.Add(durationType);
            }

        }

        /// <summary>
        /// Method to bind Time drop down
        /// </summary>
        /// <param name="comboBox">combo box</param>        
        /// <param name="addEmpty">true to add empty item</param>
        public static void BindTimeCombo(ComboBox comboBox, bool addEmpty)
        {
            List<Time> timeList = new List<Time>();
            comboBox.Items.Clear();

            if (addEmpty)
            {
                timeList.Add(new Time(-1, string.Empty));
            }


            DateTime datetime = Convert.ToDateTime(DateTime.Now.ToShortDateString());

            for (int i = 0; i < 24; i++)
            {

                timeList.Add(new Time(i, datetime.ToShortTimeString()));
                datetime = datetime.AddHours(1);

            }


            comboBox.DisplayMember = "Text";
            comboBox.ValueMember = "ID";
            comboBox.DataSource = timeList;
        }

        /// <summary>
        /// Get task flag image
        /// </summary>
        /// <param name="flag">task flag</param>
        /// <returns>flag image</returns>
        public static Image GetTaskFalgImage(TaskFlag flag)
        {
            //switch (flag)
            //{
            //    case TaskFlag.Today:
            //        {

            //            return AcmeSolar.Properties.Resources.Today_Flag;
            //        }
            //    case TaskFlag.Tomorrow:
            //        {
            //            return AcmeSolar.Properties.Resources.Tomorrow_Flag;
            //        }
            //    case TaskFlag.ThisWeek:
            //        {
            //            return AcmeSolar.Properties.Resources.ThisWeek_Flag;
            //        }
            //    case TaskFlag.NextWeek:
            //        {
            //            return AcmeSolar.Properties.Resources.NextWeek_Flag;
            //        }
            //    case TaskFlag.NoDate:
            //        {
            //            return AcmeSolar.Properties.Resources.NoDate_Flag;
            //        }

            //}


            return null;
        }

        /// <summary>
        /// get priority image
        /// </summary>
        /// <param name="priority">task priority</param>
        /// <returns>priority image</returns>
        public static Image GetTaskPriorityImage(TaskPriority priority)
        {
            //switch (priority)
            //{
            //    case TaskPriority.High:
            //        {

            //            return Properties.Resources.High_Priority;
            //        }
            //    case TaskPriority.Low:
            //        {
            //            return Properties.Resources.Low_Priority;
            //        }
            //    case TaskPriority.Medium:
            //        {
            //            return Properties.Resources.Medium_Priority;
            //        }


            //}


            return null;
        }

        /// <summary>
        /// Get  category icon
        /// </summary>
        /// <param name="category">task category</param>
        /// <returns>category icon</returns>
        public static Icon GetCategoryIcon(TaskCategory category)
        {
            //switch (category)
            //{
            //    case TaskCategory.Blue:
            //        {

            //            return Properties.Resources.Blue_Category;
            //        }
            //    case TaskCategory.Red:
            //        {

            //            return Properties.Resources.Red_Category;
            //        }
            //    case TaskCategory.Yellow:
            //        {

            //            return Properties.Resources.Yellow_Category;
            //        }
            //    case TaskCategory.Green:
            //        {

            //            return Properties.Resources.Green_Category;
            //        }
            //    case TaskCategory.Purple:
            //        {

            //            return Properties.Resources.Purple_Category;
            //        }
            //    case TaskCategory.Orange:
            //        {

            //            return Properties.Resources.Orange_Category;
            //        }

            //}


            return null;
        }

        /// <summary>
        /// Get category image
        /// </summary>
        /// <param name="category">task category</param>
        /// <returns>category image</returns>
        public static Image GetCategoryImage(TaskCategory category)
        {
            //switch (category)
            //{
            //    case TaskCategory.Blue:
            //        {

            //            return Properties.Resources.BlueCat;
            //        }
            //    case TaskCategory.Red:
            //        {

            //            return Properties.Resources.RedCat;
            //        }
            //    case TaskCategory.Yellow:
            //        {

            //            return Properties.Resources.YellowCat;
            //        }
            //    case TaskCategory.Green:
            //        {

            //            return Properties.Resources.GreenCat;
            //        }
            //    case TaskCategory.Purple:
            //        {

            //            return Properties.Resources.PurpleCat;
            //        }
            //    case TaskCategory.Orange:
            //        {

            //            return Properties.Resources.OrangeCat;
            //        }

            //}


            return null;
        }

        /// <summary>
        /// Show error message dialog and set icon in task bar
        /// </summary>
        /// <param name="message">Error message</param>
        /// <param name="caption">Error box caption</param>
        public static void ShowErrorMessage(string message, string caption)
        {
            
            MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);

        }

        /// <summary>
        /// Show confirmation dialog and set icon in task bar
        /// </summary>
        /// <param name="message">question message</param>
        /// <param name="caption">cation</param>
        /// <returns>dialog reslur</returns>
        public static DialogResult ShowConfirmationBox(string message, string caption)
        {

            DialogResult result = MessageBox.Show(message, caption, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);


            return result;
        }

        /// <summary>
        /// Draw gradient rectangle
        /// </summary>
        /// <param name="g">graphics</param>
        /// <param name="rectangle">bounds</param>
        /// <param name="color">color</param>
        public static void DrawGradientRectangle(Graphics g, Rectangle rectangle, Color color)
        {
            Color c1 = color;
            Color c2 = Color.FromArgb(AddBytes(c1.R, -14), AddBytes(c1.G, -10), AddBytes(c1.B, -5));
            Color c4 = Color.FromArgb(AddBytes(c1.R, -20), AddBytes(c1.G, -14), AddBytes(c1.B, -7));
            Color c3 = Color.FromArgb(AddBytes(c4.R, -14), AddBytes(c4.G, -10), AddBytes(c4.B, -5));

            GlossyRect(g, rectangle, c1, c2, c3, c4);

        }

        #endregion

        #region Events

        static void colorCategoryCombo_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index >= 0)
            {
                ComboBox cmbColorCategory = sender as ComboBox;
                string text = cmbColorCategory.Items[e.Index].ToString();
                e.DrawBackground();

                if (text.Length > 0)
                {
                    TaskCategory category = (TaskCategory)Enum.Parse(typeof(TaskCategory), text);
                    Icon img = GetCategoryIcon(category);
                    //Image img = Properties.Resources.YellowCat;
                    if (img != null)
                    {
                        //e.Graphics.DrawImage(img, e.Bounds.X, e.Bounds.Y, 15, 15);
                        Rectangle rect = new Rectangle(e.Bounds.X, e.Bounds.Y, 15, 15);
                        e.Graphics.DrawIcon(img, rect);

                    }
                }

                e.Graphics.DrawString(text, cmbColorCategory.Font,
                    System.Drawing.Brushes.Black,
                    new RectangleF(e.Bounds.X + 15, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));

                e.DrawFocusRectangle();
                
            }
        }

        static void priorityCombo_DrawItem(object sender, DrawItemEventArgs e)
        {

            if (e.Index >= 0)
            {
                ComboBox cmbPriority = sender as ComboBox;
                string text = cmbPriority.Items[e.Index].ToString();
                e.DrawBackground();

                if (text.Length > 0)
                {
                    TaskPriority priority = (TaskPriority)Enum.Parse(typeof(TaskPriority), text);
                    Image img = GetTaskPriorityImage(priority);

                    if (img != null)
                    {
                        e.Graphics.DrawImage(img, e.Bounds.X, e.Bounds.Y, 15, 15);

                    }
                }

                e.Graphics.DrawString(text, cmbPriority.Font,
                    System.Drawing.Brushes.Black,
                    new RectangleF(e.Bounds.X + 15, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));

                e.DrawFocusRectangle();
            }
        }

        static void taskFlagCombo_DrawItem(object sender, DrawItemEventArgs e)
        {

            if (e.Index >= 0)
            {
                ComboBox cmbTaskFlag = sender as ComboBox;
                string text = cmbTaskFlag.Items[e.Index].ToString();
                e.DrawBackground();

                if (text.Length > 0)
                {
                    TaskFlag taskFlag = (TaskFlag)Enum.Parse(typeof(TaskFlag), text);
                    Image img = Utility.GetTaskFalgImage(taskFlag);
                    //Image img = Properties.Resources.YellowCat;
                    if (img != null)
                    {
                        e.Graphics.DrawImage(img, e.Bounds.X, e.Bounds.Y, 15, 15);
                        //Rectangle rect = new Rectangle(e.Bounds.X, e.Bounds.Y, 15, 15);
                        //e.Graphics.DrawIcon(img, rect);
                    }
                }

                e.Graphics.DrawString(text, cmbTaskFlag.Font,
                    System.Drawing.Brushes.Black,
                    new RectangleF(e.Bounds.X + 15, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));

                e.DrawFocusRectangle();
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Add int value in byte
        /// </summary>
        /// <param name="b1">byte 1</param>
        /// <param name="add">added value</param>
        /// <returns>new value </returns>
        private static int AddBytes(byte b1, int add)
        {
            byte b = b1;

            try
            {

                int n = b1 + add;

                if (n > 0 && n < 256)
                    b = (byte)n;
            }
            catch
            {
            }
            return b;
        }

        /// <summary>
        /// draw glossy rectangle
        /// </summary>
        /// <param name="g">graphics</param>
        /// <param name="bounds">bounds</param>
        /// <param name="a">color a</param>
        /// <param name="b">color b</param>
        /// <param name="c">color c</param>
        /// <param name="d">color d</param>
        private static void GlossyRect(Graphics g, Rectangle bounds, Color a, Color b, Color c, Color d)
        {
            Rectangle top = new Rectangle(bounds.Left, bounds.Top, bounds.Width, bounds.Height / 2);
            Rectangle bot = Rectangle.FromLTRB(bounds.Left, top.Bottom, bounds.Right, bounds.Bottom);

            GradientRect(g, top, a, b);
            GradientRect(g, bot, c, d);

        }

        /// <summary>
        /// Draw gradient rectangle
        /// </summary>
        /// <param name="g">graphics</param>
        /// <param name="bounds">bounds</param>
        /// <param name="a">color a</param>
        /// <param name="b">color b</param>
        private static void GradientRect(Graphics g, Rectangle bounds, Color a, Color b)
        {
            if (bounds.Width > 0 && bounds.Height > 0)
            {
                using (LinearGradientBrush br = new LinearGradientBrush(bounds, b, a, -90))
                {
                    g.FillRectangle(br, bounds);
                }
            }
        }
        
        #endregion

    }


    /// <summary>
    /// 
    /// </summary>
    struct Time
    {

        #region Data Members

        int _id;
        string _text;
        
        #endregion

        #region Constructor

        public Time(int id)
            : this(id, string.Empty)
        {
        }
        public Time(int id, string text)
        {
            _id = id;
            _text = text;
        }
        
        #endregion

        #region Properties

        public string Text
        {
            get { return _text; }
            set { _text = value; }
        }

        public int ID
        {
            get { return _id; }
            set { _id = value; }
        }
        
        #endregion

    }
}
