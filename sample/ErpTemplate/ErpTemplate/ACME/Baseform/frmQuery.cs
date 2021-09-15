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
    public partial class frmQuery : Form
    {
        public frmQuery()
        {
            InitializeComponent();
        }


        private string _SqlScript;

        public string SqlScript
        {
            get
            {
                return _SqlScript;
            }

            set
            {
                 _SqlScript=value;
            }
        }

        public QuickWhere QW;

        //執行
        protected virtual void ProcessOK()
        {
            
            QW.Clear();
            //只有欄位
            //QW.Add("EMP_NO", WhereConditions.BeginsWith, TypeOfValues.StringType, textBox1);
            //QW.Add("EMP_NO", WhereConditions.BeginsWith, TypeOfValues.StringType, textBox2);

            //有 Table 的用法
            // QW.Add("EMP_NO", "RMA_EMP", TypeOfValues.StringType, WhereConditions.BeginsWith, textBox1, null);
            // QW.Add("EMP_NO", "RMA_EMP", TypeOfValues.StringType, WhereConditions.EqualTo, textBox2, null);
            SqlScript = QW.GetSql();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            ProcessOK();
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmQuery_Load(object sender, EventArgs e)
        {
            QW = new QuickWhere();
            //加入字元
            QuickWhere.SetGenerals('\'', '#', '?', '%', ',', '\\', "@@");

        }
    }
}