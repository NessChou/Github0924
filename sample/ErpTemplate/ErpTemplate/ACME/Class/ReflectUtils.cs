using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Data;
//http://info.codepub.com/2008/07/info-20320.html
namespace ACME
{
    class ReflectUtils
    {


        public static void BindAllData(Control container, Object row, DataRow dr)
        {

            Type controlType = row.GetType();
            PropertyInfo[] controlPropertiesArray = controlType.GetProperties();


            foreach (PropertyInfo controlProperty in controlPropertiesArray)
            {
                string FieldName = controlProperty.Name;

                try
                {
                    ((TextBox)FindControl(container, "TextBox" + FieldName)).Text = dr[FieldName].ToString();


                    //Binding  DateTimePicker
                    ((DateTimePicker)FindControl(container, "dp" + FieldName)).Value = ValidateUtils.StrToDate(dr[FieldName].ToString());
                    

                }
                catch
                {
                }

            }

        }

        public static void GetAllData(Control container, Object row)
        {

            Type controlType = row.GetType(); //取出所有屬性 
            PropertyInfo[] controlPropertiesArray = controlType.GetProperties();


            foreach (PropertyInfo controlProperty in controlPropertiesArray)
            {
                string FieldName = controlProperty.Name;
                SetFieldValue_TextBox(container,row, FieldName);

            }

        }

        public static void SetFieldValue_TextBox(Control container, Object row, string FieldName)
        {

            string bValue = "";
            if (((TextBox)FindControl(container, "TextBox" + FieldName)) != null)
            {
                bValue = ((TextBox)FindControl(container, "TextBox" + FieldName)).Text;
            }

            Type controlType = row.GetType(); //取出所有屬性 
            PropertyInfo[] controlPropertiesArray = controlType.GetProperties();


            foreach (PropertyInfo controlProperty in controlPropertiesArray)
            {
                if (controlProperty.Name == FieldName)
                {
                    //controlProperty.SetValue(row, bValue, null);

                    try
                    {
                        controlProperty.SetValue(row, Convert.ChangeType(bValue, controlProperty.PropertyType), null);
                    }
                    catch
                    { 
                       //Int32 空字串 一定會預設給  0
                        //controlProperty.SetValue(row, null, null);
                    }
                    break;
                }

            }

        }

        public static Control FindControl(Control container, string controlName)
        {
            Control findControl = null;
            foreach (Control control in container.Controls)
            {
                if (control.Controls.Count == 0)
                {
                    if (control.Name == controlName)
                    {
                        findControl = control;
                        break;
                    }
                }
                else
                {
                    findControl = FindControl(control, controlName);
                }
            }
            return findControl;
        }
    }
}
