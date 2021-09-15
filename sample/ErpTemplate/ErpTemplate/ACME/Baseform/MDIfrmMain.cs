using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;
using System.Xml;
using Microsoft.VisualBasic.Devices;
using System.Collections;
namespace ACME
{

    public partial class MDIfrmMain : Form
    {
        String strCn = "Data Source=10.10.1.61;Initial Catalog=acmesqlsp;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        // misc
        private bool m_ControlsLocked = false;




        public MDIfrmMain()
        {
            InitializeComponent();

            //WindowMenu.MdiList = true;

        }





        private void 水平ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void 垂直ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.TileVertical);
        }

        private void 排列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.Cascade);
        }

        private void 圖示ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void 居中ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Form activeForm = this.ActiveMdiChild;

            if (activeForm == null || activeForm.Parent == null)
                return;

            if (activeForm.WindowState != System.Windows.Forms.FormWindowState.Normal)
                return;

            activeForm.Left = (activeForm.Parent.ClientSize.Width - activeForm.Width) / 2;
            activeForm.Top = (activeForm.Parent.ClientSize.Height - activeForm.Height) / 2;

        }

        private void 關閉當前視窗ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.ActiveMdiChild != null)
                this.ActiveMdiChild.Close();
        }

        private void 關閉所有視窗ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form form in this.MdiChildren)
                form.Close();
        }


        internal void OpenMdiChildForm(Type formType)
        {
            LockControls(true);

            foreach (Form childForm in this.MdiChildren)
            {
                if (childForm.GetType() == formType)
                {
                    childForm.Activate();
                    return;
                }
            }

            Form form = Activator.CreateInstance(formType) as Form;
            if (form != null)
            {
                form.MdiParent = this;
                form.Show();
            }
            LockControls(false);
        }

        //利用名稱開啟畫面
        internal void OpenMdiChildForm(string formType)
        {
            LockControls(true);

            foreach (Form childForm in this.MdiChildren)
            {
                if ("YYJXC." + childForm.GetType().Name == formType)
                {
                    childForm.Activate();
                    LockControls(false);
                    return;
                }
            }



         

            string assemblyInformation = ", YYJXC, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null";
            Type ty = Type.GetType(formType + assemblyInformation);

            Form form = Activator.CreateInstance(ty) as Form;
            if (form != null)
            {
                form.MdiParent = this;
                form.WindowState = FormWindowState.Maximized;
                form.Show();
            }
            LockControls(false);
        }

        private void 記事本ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("notepad.exe");
        }

        private void 計算機ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("calc.exe");
        }

        private void 結束XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }


        private void LockControls(bool lockValue)
        {
            m_ControlsLocked = lockValue;

            if (m_ControlsLocked)
            {
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();
            }
            else
            {
                this.Cursor = Cursors.Arrow;
                Application.DoEvents();
            }
        }






        private void MDIfrmMain_Load(object sender, EventArgs e)
        {
       
            System.Data.DataTable dt3 = GetOrdr3(fmLogin.LoginID.ToString());
            string GG1 = "";
            string ACC = "";
            string DRS = "";
            if (dt3.Rows.Count > 0)
            {
                DataRow drw = dt3.Rows[0];
                GG1 = drw["category"].ToString();
                ACC = drw["ACC"].ToString().Trim();
                DRS = drw["DRS"].ToString().Trim();
            }


             //換資料庫
             //重新讀取
             //globals_SAP.xmlDoc = new XmlDocument();
             //string strFileName = AppDomain.CurrentDomain.BaseDirectory.ToString() + "ACME.exe.config";
             //globals_SAP.xmlDoc.Load(strFileName);
             //globals_SAP.DbName = globals_SAP.GetKeyValue("DbName");

            //共用變數
            //公司別






            if (globals.DBNAME =="進金生")
            {
                globals.shipConnectionString = "Data Source=10.10.1.61;Initial Catalog=AcmeSql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas;Connection Timeout=200000";
                globals.EEPConnectionString = "Data Source=10.10.1.61;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas;Connection Timeout=200000";
                globals.ConnectionString = "Data Source=10.10.1.61;Initial Catalog=AcmeSqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSql02";

                globals.CHOICEConnectionString  = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

            }

    
            if (globals.DBNAME == "達睿生")
            {

                globals.shipConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSql05;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSPDRS;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSqlSPDRS";
            }
            if (globals.DBNAME == "達利能源")
            {

                globals.shipConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSql09;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSql09";
            }
            if (globals.DBNAME == "進利能源")
            {

                globals.shipConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSql10;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSql10";
            }

            if (globals.DBNAME == "測試區98")
            {

                globals.shipConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSql98";
            }
            if (globals.DBNAME == "CHOICE")
            {

                globals.shipConnectionString = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSPCHOICE;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSqlSPCHOICE";
            }
            if (globals.DBNAME == "INFINITE")
            {

                globals.shipConnectionString = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSPINFINITE;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSqlSPINFINITE";
            }
            if (globals.DBNAME == "TOP GARDEN")
            {

                globals.shipConnectionString = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSPTOPGARDEN;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSqlSPTOPGARDEN";
            }
            if (globals.DBNAME == "宇豐")
            {

                globals.shipConnectionString = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSPAD;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSqlSPAD";
            }
            if (globals.DBNAME == "船務測試區")
            {
                globals.EEPConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas;Connection Timeout=200000";
                globals.shipConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSql98";
            }
            if (globals.DBNAME == "禾中")
            {

                globals.shipConnectionString = "Data Source=10.10.1.40;Initial Catalog=CHIComp23;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                globals.ConnectionString = "Data Source=acmesap;Initial Catalog=AcmeSqlSPALL;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                globals.SERVER = "AcmeSqlSPAD";
            }
            globals.shipConnection = new SqlConnection(globals.shipConnectionString);
            globals.Connection = new SqlConnection(globals.ConnectionString);
            globals.CHOICEConnection = new SqlConnection(globals.CHOICEConnectionString);
            globals.EEPConnection = new SqlConnection(globals.EEPConnectionString);

            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
  
            string NewFileName = lsAppDir + "\\ACME.exe";
            FileInfo filess = new FileInfo(NewFileName);
            string FileDate = filess.LastWriteTime.ToString();
            Text = "SAP輔助系統 登入者-" + fmLogin.LoginID + " 登入公司-" + globals.DBNAME + " 版本-" + FileDate; 

            //開啟 Logo Splash 
            ShowSplashForm();

            //選單全部資料
            //  DataTable dtMenu = GetUSERMENUS();
            //選單BY使用者權限
            string aa = GG1.Trim();
            string[] arrurl = aa.Split(new Char[] { ',' });
            StringBuilder sb = new StringBuilder();
            foreach (string i in arrurl)
            {
                sb.Append("'" + i + "',");
            }
            sb.Remove(sb.Length - 1, 1);

            int GG2 = aa.IndexOf(",");
            if (GG2 != -1)
            {
                globals.GroupID = aa.Substring(0, GG2);
            }
            else
            {
                globals.GroupID = aa;
            }

            DataTable dtMenu = GetUSERMENUS(sb.ToString(), ACC,DRS);


            //載入選單
            //menuStrip1.Items.Clear();

            ToolStripMenuItem menuItem = new ToolStripMenuItem();

            menuItem.Name = "System";
            menuItem.Text = "模組(&M)";
            menuStrip1.Items.Insert(1, menuItem); //
            //Add(menuItem);

            GenerateMenus(dtMenu, "", menuItem);

      

        }

        private void MenuItemOnClick_Open(object sender, EventArgs e)
        {
            //  MessageBox.Show("Open Clicked");
            //取得代號
            // MessageBox.Show(((ToolStripMenuItem)sender).Name);
            try
            {
                string FormName = "ACME." + ((ToolStripMenuItem)sender).Name;
                string AA = ((ToolStripMenuItem)sender).Text;
                object aObject = Activator.CreateInstance(Type.GetType(FormName));
                Form aForm = aObject as Form;
                aForm.WindowState = FormWindowState.Maximized;
                aForm.MdiParent = this;
                aForm.Show();

                if (fmLogin.LoginID.ToString().ToUpper().Trim() != "LLEYTONCHEN")
                {
                    if (globals.DBNAME == "進金生")
                    {
                        GetMenu.InsertLog(fmLogin.LoginID.ToString(), "LOGIN", AA, DateTime.Now.ToString("yyyyMMddHHmmss"));
                    }
                }

            }
            catch (Exception ex)
            {
                //if (fmLogin.LoginID.ToString().ToUpper().Trim() != "SANDYLO")
                //{
                //    MessageBox.Show("Error->" + ex.Message);
                //}
            }
        }

        private void ShowSplashForm()
        {

            //開啟 Logo Splash 
            SplashForm sForm = new SplashForm();
            sForm.Show();
            sForm.Refresh();
            System.Threading.Thread.Sleep(1000);
            sForm.Close();
        }

        //處理系統選單
        //如果是 Parnet =""; 則不需要連結事件 20071028
        private void GenerateMenus(DataTable dt, string rootID, ToolStripMenuItem menuItem)
        {

            ToolStripItem item = null;
            ToolStripSeparator separator = null;


            string strExp;

            if (rootID == "")
            {
                strExp = "[PARENT] is null or [PARENT]=''";
            }
            else
            {
                strExp = "[PARENT] = '" + rootID + "'";
            }

            DataRow[] childRows = dt.Select(strExp);

            foreach (DataRow dr in childRows)
            {

                string rowID = dr["MENUID"].ToString();


                string Parent = dr["PARENT"].ToString();



                string dept = Convert.ToString(dr["MENUID"]);



                if (menuItem == null)
                {
                    menuItem = new ToolStripMenuItem();
                    menuItem.Name = Convert.ToString(dr["MENUID"]); ;
                    menuItem.Text = Convert.ToString(dr["CAPTION"]);
                    menuStrip1.Items.Add(menuItem);
                }
                else
                {
                    item = new ToolStripMenuItem();
                    item.Name = Convert.ToString(dr["MENUID"]);
                    item.Text = Convert.ToString(dr["CAPTION"]);

                    menuItem.DropDownItems.Add(item);

                    //20071028
                    if (!string.IsNullOrEmpty(Parent))
                    {
                        FindEventsByName(item, this, true, "MenuItemOn", "_Open");
                    }
                }


                GenerateMenus(dt, rowID, (ToolStripMenuItem)item);

            }

        }

        private void FindEventsByName(object sender, object receiver, bool bind, string handlerPrefix, string handlerSuffix)
        {
            System.Reflection.EventInfo[] SenderEvent = sender.GetType().GetEvents();
            Type ReceiverType = receiver.GetType();
            System.Reflection.MethodInfo method;

            foreach (System.Reflection.EventInfo e in SenderEvent)
            {
                method = ReceiverType.GetMethod(string.Format("{0}{1}{2}", handlerPrefix, e.Name, handlerSuffix), System.Reflection.BindingFlags.IgnoreCase | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);

                if (method != null)
                {

                    System.Delegate d = System.Delegate.CreateDelegate(e.EventHandlerType, receiver, method.Name);

                    if (bind)
                        e.AddEventHandler(sender, d);
                    else
                        e.RemoveEventHandler(sender, d);
                }
            }
        }



        private DataTable GetUSERMENUS(string USERID, string ACC, string DRS)
        {
            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            if (!String.IsNullOrEmpty(ACC))
            {

                if (ACC == "601")
                {
                    if (globals.DBNAME == "達睿生")
                    {
                        sb.Append(" SELECT DISTINCT M.PARENT,M.MENUID,M.CAPTION FROM MENUTABLE2 M WHERE (M.MENUID = '601' OR M.PARENT='601')");
                        sb.Append(" UNION ALL");
                        sb.Append(" SELECT DISTINCT M.PARENT,M.MENUID,M.CAPTION FROM MENUTABLE2 M WHERE (M.MENUID = '203' OR M.PARENT='203')");
                        sb.Append(" UNION ALL");
                    }
                    else
                    {
                        sb.Append(" SELECT DISTINCT M.PARENT,M.MENUID,M.CAPTION FROM MENUTABLE2 M WHERE (M.MENUID <> '203')");
                        sb.Append(" UNION ALL");
                    }
                }
                else if (ACC == "203")
                {
                    sb.Append(" SELECT DISTINCT M.PARENT,M.MENUID,M.CAPTION FROM MENUTABLE2 M WHERE (M.MENUID = '203' OR M.PARENT='203')");
                    sb.Append(" UNION ALL");
                }
                else if (ACC == "609")
                {
                    sb.Append(" SELECT DISTINCT M.PARENT,M.MENUID,M.CAPTION FROM MENUTABLE2 M");
                    sb.Append(" WHERE (M.MENUID NOT IN ('601','203','604','605'))");
                    sb.Append(" UNION ALL");
                }
                else
                {
                    sb.Append(" SELECT DISTINCT M.PARENT,M.MENUID,M.CAPTION FROM MENUTABLE2 M");
                    sb.Append(" WHERE (M.MENUID NOT IN ('601','203'))");
                    sb.Append(" UNION ALL");
                }
            }
            sb.Append(" SELECT DISTINCT M.PARENT,M.MENUID,M.CAPTION FROM USERMENUS U ");
                      sb.Append(" INNER JOIN MENUTABLE M ON M.MENUID=U.MENUID  ");
                         sb.Append(" WHERE U.USERID IN (" + USERID + "  )  ");
                         if (globals.DBNAME == "進金生")
                         {
                             sb.Append(" AND  U.MENUID <> 'DRSTT' ");
                         }
                         if (globals.DBNAME == "達睿生")
                         {
                             sb.Append(" AND  U.MENUID <> 'TT' ");
                         }
                         if (globals.DBNAME == "宇豐")
                         {
                             sb.Append(" AND  (  U.ENABLED ='Y' OR U.MENUID='RMA19') ");
                         }
                         else  if (globals.DBNAME == "禾中")
                         {
                             sb.Append(" AND  U.ENABLED ='Y'   AND U.MENUID NOT IN ('fmShip','CAR','CHECKSHIP','LAB','PLATE','RmaCarton','WH_ITEM1') ");
                         }
                         else
                         {
                             sb.Append(" AND  U.ENABLED ='Y' AND U.MENUID<>'RMA20' ");
                         }

                         if (DRS == "V")
                         {
                             sb.Append(" UNION ALL");
                             sb.Append(" SELECT  'RMA21' PARENT,M.MENUID,M.CAPTION  FROM MENUTABLE U　INNER JOIN MENUTABLE M ON M.MENUID=U.MENUID 　WHERE U.MENUID IN ('DOCUPLOADS','SHICAR','SHIPMULTI','SHIPPACK')");
                         }

                         SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "MENUS");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["MENUS"];
        }


        public DataTable GetOrdr3(string username)
        {
            SqlConnection connection = new SqlConnection(strCn);

            string sql = "select category,ACC,ChineseName DRS from [right] where username=@username";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@username", username));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "right");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["right"];
        }

        private void 更新程式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UPLOAD();
        }
        private void UPLOAD()
        {
            DialogResult result;
            result = MessageBox.Show("將關閉所有輔助程式，請先存檔", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string NewFileName = lsAppDir + "\\更新程式.exe";

                ArrayList ar = GetFileList(lsAppDir);
                foreach (string Fil in ar)
                {
                    FileInfo info = new FileInfo(Fil);
                    string G1 = info.Name.ToString();
                    int T2 = G1.IndexOf("更新程式");
                    if (T2 != -1)
                    {
                        System.IO.File.Delete(Fil);
                    }
                }

                string OutPutFile = GetPATH().Rows[0][0].ToString() + "\\更新程式.exe";
                System.IO.File.Copy(OutPutFile, NewFileName, true);


                System.Diagnostics.Process.Start(NewFileName);
                Close();
            }
        }


        public DataTable GetPATH()
        {
            SqlConnection connection = new SqlConnection(strCn);

            string sql = "SELECT PARAM_NO  FROM RMA_PARAMS WHERE PARAM_KIND='COPYPATH2'";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "right");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["right"];
        }
        private static ArrayList GetFileList(string Dir)
        {
            ArrayList fils = new ArrayList();
            bool Empty = true;
            foreach (string file in Directory.GetFiles(Dir))
            {
                fils.Add(file);
                Empty = false;
            }

            if (Empty)
            {
                if (Directory.GetDirectories(Dir).Length == 0)
                    fils.Add(Dir + @"/");

            }

            foreach (string dirs in Directory.GetDirectories(Dir))
            {
                foreach (object obj in GetFileList(dirs))
                {
                    fils.Add(obj);
                }
            }
            return fils;
        }

    

        private void timer1_Tick(object sender, EventArgs e)
        {
                  System.Data.DataTable G1 = GETORTT();
                  if (G1.Rows.Count == 0)
                  {
                      HRTEMP frm1 = new HRTEMP();
                      frm1.Show();
                  }
        }
        private System.Data.DataTable GETORTT()
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT YTEMP FROM HR_TEMP WHERE DOCDATE=@DOCDATE AND USERS=@USERS");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCDATE", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void 量體溫ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //HRTEMP frm1 = new HRTEMP();
            //frm1.Show();
        }

        private void menuStrip1_Click(object sender, EventArgs e)
        {
            //if (fmLogin.LoginID.ToString().ToUpper().Trim() == "LLEYTONCHEN" || fmLogin.LoginID.ToString().ToUpper().Trim() == "FEDERLIU")
            //{
            //    System.Data.DataTable G1 = GETORTT();
            //    if (G1.Rows.Count == 0)
            //    {
            //        HRTEMP frm1 = new HRTEMP();
            //        frm1.Show();

            //        //  MessageBox.Show("請至EIP填寫今日體溫http://www.google.com");

            //    }
            //}
        }

  


    }
}