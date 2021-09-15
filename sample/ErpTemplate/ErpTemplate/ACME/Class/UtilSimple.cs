using System;
using System.Collections.Generic;
using System.Text;
using System.Security.Cryptography;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;


namespace ACME
{
    class UtilSimple
    {


    

        public static string Encrypt(string clearText, string Password)
        {
            byte[] clearBytes =
                Encoding.Unicode.GetBytes(clearText);

            //PasswordDeriveBytes pdb = new PasswordDeriveBytes(Password,
            //    new byte[] {0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d,
            //                       0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76});
            PasswordDeriveBytes pdb = new PasswordDeriveBytes(Password,
                new byte[] { 0x49 });

            byte[] encryptedData = Encrypt(clearBytes,
                pdb.GetBytes(32), pdb.GetBytes(16));

            return Convert.ToBase64String(encryptedData);

        }

        public static byte[] Encrypt(byte[] clearData, byte[] Key, byte[] IV)
        {
            MemoryStream ms = new MemoryStream();

            Rijndael alg = Rijndael.Create();


            alg.Key = Key;
            alg.IV = IV;

            CryptoStream cs = new CryptoStream(ms,
                alg.CreateEncryptor(), CryptoStreamMode.Write);

            cs.Write(clearData, 0, clearData.Length);

            cs.Close();

            byte[] encryptedData = ms.ToArray();

            return encryptedData;
        }

        public static byte[] Decrypt(byte[] cipherData,
                byte[] Key, byte[] IV)
        {
            MemoryStream ms = new MemoryStream();

            Rijndael alg = Rijndael.Create();

            alg.Key = Key;
            alg.IV = IV;

            CryptoStream cs = new CryptoStream(ms,
                alg.CreateDecryptor(), CryptoStreamMode.Write);

            cs.Write(cipherData, 0, cipherData.Length);

            cs.Close();

            byte[] decryptedData = ms.ToArray();

            return decryptedData;
        }

        public static string Decrypt(string cipherText, string Password)
        {
            byte[] cipherBytes = Convert.FromBase64String(cipherText);

            //PasswordDeriveBytes pdb = new PasswordDeriveBytes(Password,
            //    new byte[] {0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65,
            //                       0x64, 0x76, 0x65, 0x64, 0x65, 0x76});

            PasswordDeriveBytes pdb = new PasswordDeriveBytes(Password,
                new byte[] { 0x49 });

            byte[] decryptedData = Decrypt(cipherBytes,
                pdb.GetBytes(32), pdb.GetBytes(16));

            return Encoding.Unicode.GetString(decryptedData);
        }


        public static string PADRight(string str, int len)
        {
            string s = "0000000000" + str;
            return s.Substring(s.Length - len, len);
        }

        public static string GetMaxID(SqlConnection MyConnection,  string MyTableName)
        {
            string sSQL = "SELECT MAX(ShippingCode)+1 AS ID FROM " + MyTableName;
            SqlCommand cmdSQL = new SqlCommand();
            cmdSQL.CommandText = sSQL;
            cmdSQL.Connection = MyConnection;
            MyConnection.Open();
            try
            {
                string NewID = cmdSQL.ExecuteScalar().ToString();

                return NewID;
            }
            finally
            {
                MyConnection.Close();
            }
        }

        //開啟 Lookup 視窗
        public static object[] GetMenuList()
        {
            string[] FieldNames = new string[] { "MENUID", "CAPTION" };

            string[] Captions = new string[] { "選單代號", "名稱" };

            string SqlScript = "SELECT * FROM MENUTABLE";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        //開啟 Lookup 視窗
        public static object[] GetCardList()
        {
            string[] FieldNames = new string[] { "CardCode", "CardName" };

            string[] Captions = new string[] { "代碼", "名稱" };

            string SqlScript = "SELECT CardCode,CardName FROM OCRD";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        //
        public static void SetLookupBinding(ListControl toBind, object dataSource, string displayMember, string valueMember)
        {

            //toBind.DataBindings.Clear();
            ////要加入這個才會正確
            //toBind.DataBindings.Add(new Binding("SelectedValue", this.oRDRBindingSource, "CurSource", true));
            
            
            toBind.DisplayMember = displayMember;
            toBind.ValueMember = valueMember;

            // Only after the following line will the listbox receive events due to binding.
            toBind.DataSource = dataSource;
        }


        public static void SetLookupBinding(ListControl toBind,string Param_kind)
        {

            //toBind.DataBindings.Clear();
            ////要加入這個才會正確
            //toBind.DataBindings.Add(new Binding("SelectedValue", this.oRDRBindingSource, "CurSource", true));


            toBind.DisplayMember = "PARAM_DESC";
            toBind.ValueMember = "PARAM_NO";

            toBind.DataSource = LookupCode(Param_kind);

            
        }

        public static void SetLookupBinding(ListControl toBind, string Param_kind,BindingSource BS,string FieldName)
        {

            //toBind.DataBindings.Clear();
            ////要加入這個才會正確
            //toBind.DataBindings.Add(new Binding("SelectedValue", this.oRDRBindingSource, "CurSource", true));


            toBind.DisplayMember = "PARAM_DESC";
            toBind.ValueMember = "PARAM_NO";

            toBind.DataSource = LookupCode(Param_kind);
            toBind.DataBindings.Clear();
            toBind.DataBindings.Add(new Binding("SelectedValue", BS, FieldName, true));
        }


        //傳入自訂查詢的資料表
        public static void SetLookupBinding(ListControl toBind,SqlConnection Connection,  string Sql, BindingSource BS, string FieldName,
            string displayMember, string valueMember)
        {

            toBind.DisplayMember = displayMember;
            toBind.ValueMember = valueMember;

            toBind.DataSource = GetDataTable(Connection,Sql);
            toBind.DataBindings.Clear();
            toBind.DataBindings.Add(new Binding("SelectedValue", BS, FieldName, true));
        }



        //傳入自訂查詢的資料表
        public static void SetLookupBinding(DataGridViewComboBoxColumn toBind, SqlConnection Connection, string Sql, 
            string displayMember, string valueMember)
        {

            toBind.DisplayMember = displayMember;
            toBind.ValueMember = valueMember;

            toBind.DataSource = GetDataTable(Connection, Sql);
        }


        public static void SetLookupBinding(DataGridViewComboBoxColumn toBind, string Param_kind)
        {

            //toBind.DataBindings.Clear();
            ////要加入這個才會正確
            //toBind.DataBindings.Add(new Binding("SelectedValue", this.oRDRBindingSource, "CurSource", true));


            toBind.DisplayMember = "PARAM_DESC";
            toBind.ValueMember = "PARAM_NO";

            toBind.DataSource = LookupCode(Param_kind);


        }


        public static DataTable LookupCode(string Param_kind)
        {
            string Sql = "select * from RMA_PARAMS where Param_kind='" + Param_kind + "'";
            DataTable dt = GetDataTable(globals.Connection, Sql);

            return dt;

        }


        public static System.Data.DataTable GetDataTable(SqlConnection MyConnection, string sql)
        {

            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "Data");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["Data"];
        }
    }
}
