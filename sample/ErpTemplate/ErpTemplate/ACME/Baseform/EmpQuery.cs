using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using DragD.QuickWhereComponent;

namespace YYJXC.BaseForm
{
    public partial class EmpQuery : YYJXC.BaseForm.frmQuery
    {
        public EmpQuery()
        {
            InitializeComponent();
        }

        //執行
        protected override void ProcessOK()
        {

            QW.Clear();
            ////只有欄位
            ////QW.Add("EMP_NO", WhereConditions.BeginsWith, TypeOfValues.StringType, textBox1);
            ////QW.Add("EMP_NO", WhereConditions.BeginsWith, TypeOfValues.StringType, textBox2);

            ////有 Table 的用法
             QW.Add("EMP_NO", "RMA_EMP", TypeOfValues.StringType, WhereConditions.BeginsWith, textBox1, null);
             QW.Add("EMP_NO", "RMA_EMP", TypeOfValues.StringType, WhereConditions.EqualTo, textBox2, null);
            SqlScript = QW.GetSql();
        }
    }
}

